"""A small Excel formula evaluator for the tests — the subset compare and action3
write (IF, OR, AND, NOT, IFERROR, TRIM, VALUE, SUBSTITUTE, ABS, ROUND, MAX, MIN,
LEN, CHAR, EXACT, ISNUMBER, ISTEXT, SEARCH, FIND, LEFT, LOWER; cell refs incl. 'Sheet'!$A$1 and ranges
$A1:$C1; & + - * comparisons), so a test can check what a spreadsheet would
SHOW instead of matching formula text.

    from xlformula import evaluate, cf_kind
    evaluate(wb, ws, '=IF(A1="","N/A",A1*2)')     # → value, '' for empty
    cf_kind(wb, ws, 'E7', kind_of)                # → the first conditional-format
                                                  #   rule that fires on E7

Errors (#VALUE! etc.) raise XLError; IFERROR catches them."""
import re
from decimal import Decimal, ROUND_HALF_UP

__all__ = ['evaluate', 'cf_kind', 'XLError']


class XLError(Exception):
    pass


class _Empty:
    """An empty cell: "" in text, 0 in arithmetic."""
    def __repr__(self):
        return 'EMPTY'


EMPTY = _Empty()

_TOKEN = re.compile(r"""
    (?P<ws>\s+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<range>\$?[A-Z]{1,3}\$?[0-9]+:\$?[A-Z]{1,3}\$?[0-9]+)
  | (?P<ref>(?:'[^']+'!|[A-Za-z_][\w ]*!)?\$?[A-Z]{1,3}\$?[0-9]+)
  | (?P<num>[0-9]+(?:\.[0-9]*)?|\.[0-9]+)
  | (?P<bool>TRUE|FALSE)(?![A-Za-z0-9_(])
  | (?P<func>[A-Z][A-Z0-9.]*)(?=\()
  | (?P<op><>|<=|>=|[-+*/&=<>(),])
""", re.X)


def _tokens(src):
    pos, out = 0, []
    while pos < len(src):
        m = _TOKEN.match(src, pos)
        if not m:
            raise ValueError(f'cannot parse at {src[pos:pos + 20]!r}')
        pos = m.end()
        kind = m.lastgroup
        if kind != 'ws':
            out.append((kind, m.group(kind)))
    return out


def _num(v):
    if v is EMPTY:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError:
        raise XLError('#VALUE!')


def _text(v):
    if v is EMPTY or v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float):
        return ('%.15g' % v)
    return str(v)


def _cmp(a, b, op):
    if a is EMPTY and isinstance(b, str):
        a = ''
    if b is EMPTY and isinstance(a, str):
        b = ''
    if a is EMPTY:
        a = 0.0
    if b is EMPTY:
        b = 0.0
    na, nb = isinstance(a, (int, float)) and not isinstance(a, bool), isinstance(b, (int, float)) and not isinstance(b, bool)
    if na and nb:
        x, y = float(a), float(b)
    elif not na and not nb:
        x, y = _text(a).lower(), _text(b).lower()
    else:                       # Excel: text sorts after numbers; never equal
        x, y = (0, 1) if na else (1, 0)
    return {'=': x == y, '<>': x != y, '<': x < y, '>': x > y, '<=': x <= y, '>=': x >= y}[op]


class _Range(list):
    """The values of a cell range (only MAX / MIN take one)."""


class _Parser:
    def __init__(self, toks, getcell, getrange=None):
        self.t, self.i, self.get, self.getrange = toks, 0, getcell, getrange

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self, val=None):
        tok = self.peek()
        if val is not None and tok[1] != val:
            raise ValueError(f'expected {val!r}, got {tok}')
        self.i += 1
        return tok

    # each parse_* returns a thunk so IF / IFERROR evaluate lazily
    def expr(self):
        left = self.concat()
        while self.peek()[1] in ('=', '<>', '<', '>', '<=', '>='):
            op = self.take()[1]
            right = self.concat()
            left = (lambda l, r, o: lambda: _cmp(l(), r(), o))(left, right, op)
        return left

    def concat(self):
        left = self.additive()
        while self.peek()[1] == '&':
            self.take()
            right = self.additive()
            left = (lambda l, r: lambda: _text(l()) + _text(r()))(left, right)
        return left

    def additive(self):
        left = self.mult()
        while self.peek()[1] in ('+', '-'):
            op = self.take()[1]
            right = self.mult()
            left = (lambda l, r, o: lambda: _num(l()) + _num(r()) if o == '+' else _num(l()) - _num(r()))(
                left, right, op)
        return left

    def mult(self):
        left = self.unary()
        while self.peek()[1] in ('*', '/'):
            op = self.take()[1]
            right = self.unary()
            left = (lambda l, r, o: lambda: _num(l()) * _num(r()) if o == '*' else _div(_num(l()), _num(r())))(
                left, right, op)
        return left

    def unary(self):
        if self.peek()[1] == '-':
            self.take()
            inner = self.unary()
            return lambda: -_num(inner())
        if self.peek()[1] == '+':
            self.take()
            return self.unary()
        return self.primary()

    def primary(self):
        kind, val = self.take()
        if kind == 'num':
            return lambda v=float(val): v
        if kind == 'bool':
            return lambda v=(val == 'TRUE'): v
        if kind == 'str':
            return lambda v=val[1:-1].replace('""', '"'): v
        if kind == 'ref':
            return lambda r=val: self.get(r)
        if kind == 'range':
            return lambda r=val: _Range(self.getrange(r))
        if val == '(':
            e = self.expr()
            self.take(')')
            return e
        if kind == 'func':
            self.take('(')
            args = []
            if self.peek()[1] != ')':
                args.append(self.expr())
                while self.peek()[1] == ',':
                    self.take()
                    args.append(self.expr())
            self.take(')')
            return _func(val, args)
        raise ValueError(f'unexpected {kind} {val!r}')


def _div(a, b):
    if b == 0:
        raise XLError('#DIV/0!')
    return a / b


def _value_of_text(s):
    """VALUE() of a text, as Excel / LibreOffice (en-US) read it: spaces around
    are fine, a tab / newline / non-ASCII digit is not; "1,000", "10%", "(5)",
    "$9.45", "9:45" (a time), "9 1/2" (a fraction) and "1/2" (a date) are numbers."""
    s = s.strip(' ')
    if not s or not s.isascii():
        raise XLError('#VALUE!')
    neg = s.startswith('(') and s.endswith(')')
    if neg:
        s = s[1:-1].strip(' ')
    pct = s.endswith('%')
    if pct:
        s = s[:-1].strip(' ')
    m = re.fullmatch(r'([+-]?)\$?\s*([0-9]{1,3}(?:,[0-9]{3})+(?:\.[0-9]*)?|[0-9]+\.?[0-9]*(?:[eE][+-]?[0-9]+)?|\.[0-9]+)', s)
    if m:
        x = float(m.group(2).replace(',', '')) * (-1 if m.group(1) == '-' else 1)
    elif re.fullmatch(r'[0-9]+:[0-9]{2}(?::[0-9]{2})?', s):
        h, mi, *sec = (int(p) for p in s.split(':'))
        x = h / 24 + mi / 1440 + (sec[0] / 86400 if sec else 0)
    elif re.fullmatch(r'[0-9]+ [0-9]+/[0-9]+', s):
        w, fr = s.split(' ')
        a_, b_ = (int(p) for p in fr.split('/'))
        if b_ == 0:
            raise XLError('#VALUE!')
        x = w and int(w) + a_ / b_
    elif re.fullmatch(r'[0-9]{1,2}/[0-9]{1,2}(?:/[0-9]{2,4})?', s):
        x = 45000.0                               # a date: some serial number (> 1000)
    else:
        raise XLError('#VALUE!')
    x = -x if neg else x
    return x / 100 if pct else x


def _func(name, args):
    if name == 'IF':
        return lambda: (args[1]() if _truth(args[0]()) else (args[2]() if len(args) > 2 else False))
    if name == 'IFERROR':
        def f():
            try:
                return args[0]()
            except XLError:
                return args[1]()
        return f
    if name == 'OR':
        return lambda: any(_truth(a()) for a in args)
    if name == 'AND':
        return lambda: all(_truth(a()) for a in args)
    if name == 'NOT':
        return lambda: not _truth(args[0]())
    if name == 'LEN':
        return lambda: float(len(_text(args[0]())))
    if name == 'CHAR':
        return lambda: chr(int(_num(args[0]())))
    if name == 'EXACT':
        return lambda: _text(args[0]()) == _text(args[1]())
    if name == 'TRIM':
        return lambda: re.sub(' +', ' ', _text(args[0]()).strip(' '))
    if name == 'LOWER':
        return lambda: _text(args[0]()).lower()
    if name == 'VALUE':
        def f():
            v = args[0]()
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return float(v)
            return _value_of_text(_text(v))
        return f
    if name == 'SUBSTITUTE':
        return lambda: _text(args[0]()).replace(_text(args[1]()), _text(args[2]()))
    if name == 'ABS':
        return lambda: abs(_num(args[0]()))
    if name == 'ROUND':
        def f():
            x, n = _num(args[0]()), int(_num(args[1]()))
            return float(Decimal(repr(x)).quantize(Decimal(1).scaleb(-n), rounding=ROUND_HALF_UP))
        return f
    if name in ('MAX', 'MIN'):
        def f():
            vals = []
            for a in args:
                v = a()
                if isinstance(v, _Range):           # a range: numbers only (text / empty skipped)
                    vals += [float(x) for x in v if isinstance(x, (int, float)) and not isinstance(x, bool)]
                elif v is not EMPTY:
                    vals.append(_num(v))
            return (max if name == 'MAX' else min)(vals) if vals else 0.0
        return f
    if name in ('ISNUMBER', 'ISTEXT'):
        def f():                                # an IS function never passes an error on
            try:
                v = args[0]()
            except XLError:
                return False
            if name == 'ISTEXT':
                return isinstance(v, str)
            return isinstance(v, (int, float)) and not isinstance(v, bool)
        return f
    if name == 'ISERROR':
        def f():
            try:
                args[0]()
                return False
            except XLError:
                return True
        return f
    if name in ('SEARCH', 'FIND'):              # FIND: case-sensitive (neither uses wildcards here)
        def f():
            within, what = _text(args[1]()), _text(args[0]())
            pos = (within.lower().find(what.lower()) if name == 'SEARCH' else within.find(what))
            if pos < 0:
                raise XLError('#VALUE!')
            return float(pos + 1)
        return f
    if name == 'LEFT':
        return lambda: _text(args[0]())[:int(_num(args[1]())) if len(args) > 1 else 1]
    raise ValueError(f'function {name} not supported')


def _truth(v):
    if isinstance(v, bool):
        return v
    if v is EMPTY:
        return False
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str) and v.upper() in ('TRUE', 'FALSE'):
        return v.upper() == 'TRUE'
    raise XLError('#VALUE!')


def evaluate(wb, ws, formula, depth=0):
    """Value of a formula (or a plain value) on sheet `ws` of workbook `wb`.
    Referenced cells that hold formulas are evaluated too. Empty result → ''."""
    if depth > 20:
        raise RecursionError('formula chain too deep')

    def getcell(ref):
        m = re.fullmatch(r"(?:'([^']+)'!|([A-Za-z_][\w ]*)!)?\$?([A-Z]{1,3})\$?([0-9]+)", ref)
        sheet = wb[m.group(1) or m.group(2)] if (m.group(1) or m.group(2)) else ws
        cell = sheet[f'{m.group(3)}{m.group(4)}']
        v = cell.value
        if isinstance(v, str) and v.startswith('=') and cell.data_type == 'f':    # text '=…' stays text
            v = evaluate(wb, sheet, v, depth + 1)
            return EMPTY if v == '' else v
        return EMPTY if v is None else v

    def getrange(ref):
        a, b = ref.replace('$', '').split(':')
        out = []
        for row in ws[a:b]:
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('=') and c.data_type == 'f':
                    v = evaluate(wb, ws, v, depth + 1)
                out.append(None if v in (None, '') else v)
        return out

    if not (isinstance(formula, str) and formula.startswith('=')):
        return '' if formula is None else formula
    p = _Parser(_tokens(formula[1:]), getcell, getrange)
    out = p.expr()()
    if p.i != len(p.t):
        raise ValueError(f'trailing tokens in {formula!r}')
    return '' if out is EMPTY else out


def _shift_rows(formula, by):
    """A conditional-format formula written for its range's first row, moved
    `by` rows down (relative row numbers only — '$7' stays)."""
    return re.sub(r'(?<![A-Z$])(\$?[A-Z]{1,3})([0-9]+)\b',
                  lambda m: f'{m.group(1)}{int(m.group(2)) + by}', formula)


def cf_kind(wb, ws, coord, kind_of):
    """kind_of(rule) of the first conditional-format rule (by priority) whose
    formula is TRUE for cell `coord` — what the spreadsheet would show — or None.
    Rules that raise count as not firing (a spreadsheet shows no format)."""
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    col_l, row = coordinate_from_string(coord)
    col = column_index_from_string(col_l)
    hits = []
    for cf in ws.conditional_formatting:
        for rng in str(cf.sqref).split():
            a, _, b = rng.partition(':')
            (ca, ra), (cb, rb) = coordinate_from_string(a), coordinate_from_string(b or a)
            if column_index_from_string(ca) <= col <= column_index_from_string(cb) and ra <= row <= rb:
                hits += [(r.priority, r, ra) for r in cf.rules]
    for _p, rule, top in sorted(hits, key=lambda h: h[0]):
        f = _shift_rows(rule.formula[0], row - top)
        try:
            v = evaluate(wb, ws, '=' + f)
        except XLError:
            continue
        if _truth_soft(v):
            return kind_of(rule)
    return None


def _truth_soft(v):
    try:
        return _truth(v)
    except XLError:
        return False
