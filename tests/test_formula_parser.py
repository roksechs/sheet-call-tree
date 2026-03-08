"""Unit tests for formula_parser — formula string → typed AST."""
import pytest

from sheet_call_tree.formula_parser import parse_formula
from sheet_call_tree.models import CellNode, FunctionNode, RangeNode


SHEET = "Sheet1"


def p(formula):
    return parse_formula(formula, SHEET)


def ref(r):
    """Shorthand: CellNode with defaults."""
    return CellNode(cell=r)


def fn(name, *args):
    """Shorthand: FunctionNode."""
    return FunctionNode(name, list(args))


def rng(start, end):
    """Shorthand: RangeNode from two ref strings."""
    return RangeNode(start, end)


class TestOperands:
    def test_single_cell_no_sheet(self):
        assert p("=A1") == ref("Sheet1!A1")

    def test_single_cell_with_sheet(self):
        assert p("=Sheet2!C5") == ref("Sheet2!C5")

    def test_single_cell_absolute(self):
        assert p("=$A$1") == ref("Sheet1!A1")

    def test_number_integer(self):
        assert p("=42") == 42

    def test_number_float(self):
        assert p("=1.1") == 1.1

    def test_string_literal(self):
        assert p('="hello"') == "hello"

    def test_logical_true(self):
        assert p("=TRUE") is True

    def test_logical_false(self):
        assert p("=FALSE") is False


class TestRanges:
    def test_range_unqualified(self):
        assert p("=SUM(A1:A9)") == fn("SUM", rng("Sheet1!A1", "Sheet1!A9"))

    def test_range_qualified(self):
        assert p("=SUM(Sheet2!A1:A9)") == fn("SUM", rng("Sheet2!A1", "Sheet2!A9"))

    def test_column_range(self):
        result = p("=VLOOKUP(A1,Sheet2!A:B,2,FALSE)")
        assert result == fn(
            "VLOOKUP",
            ref("Sheet1!A1"),
            rng("Sheet2!A", "Sheet2!B"),
            2,
            False,
        )


class TestInfixOperators:
    def test_add(self):
        assert p("=A1+1") == fn("ADD", ref("Sheet1!A1"), 1)

    def test_sub(self):
        assert p("=A1-1") == fn("SUB", ref("Sheet1!A1"), 1)

    def test_mul(self):
        assert p("=C5*2.0") == fn("MUL", ref("Sheet1!C5"), 2.0)

    def test_div(self):
        assert p("=A1/B1") == fn("DIV", ref("Sheet1!A1"), ref("Sheet1!B1"))

    def test_pow(self):
        assert p("=A1^2") == fn("POW", ref("Sheet1!A1"), 2)

    def test_concat(self):
        assert p('="hello"&A1') == fn("CONCAT", "hello", ref("Sheet1!A1"))

    def test_eq(self):
        assert p("=A1=0") == fn("EQ", ref("Sheet1!A1"), 0)

    def test_neq(self):
        assert p("=A1<>0") == fn("NEQ", ref("Sheet1!A1"), 0)

    def test_gt(self):
        assert p("=A1>0") == fn("GT", ref("Sheet1!A1"), 0)

    def test_lt(self):
        assert p("=A1<0") == fn("LT", ref("Sheet1!A1"), 0)


class TestPrecedence:
    def test_mul_before_add(self):
        assert p("=A1+B1*C1") == fn(
            "ADD",
            ref("Sheet1!A1"),
            fn("MUL", ref("Sheet1!B1"), ref("Sheet1!C1")),
        )

    def test_parens_override_precedence(self):
        assert p("=(A1+B1)*C1") == fn(
            "MUL",
            fn("ADD", ref("Sheet1!A1"), ref("Sheet1!B1")),
            ref("Sheet1!C1"),
        )

    def test_pow_right_associative(self):
        assert p("=2^3^4") == fn("POW", 2, fn("POW", 3, 4))

    def test_left_assoc_sub(self):
        assert p("=A1-B1-C1") == fn(
            "SUB",
            fn("SUB", ref("Sheet1!A1"), ref("Sheet1!B1")),
            ref("Sheet1!C1"),
        )


class TestPrefixOperators:
    def test_unary_neg(self):
        assert p("=-A1") == fn("NEG", ref("Sheet1!A1"))

    def test_unary_neg_in_expression(self):
        assert p("=2*-A1") == fn("MUL", 2, fn("NEG", ref("Sheet1!A1")))

    def test_neg_in_function_arg(self):
        assert p("=IF(A1>0,B1,-C1)") == fn(
            "IF",
            fn("GT", ref("Sheet1!A1"), 0),
            ref("Sheet1!B1"),
            fn("NEG", ref("Sheet1!C1")),
        )


class TestFunctions:
    def test_sum_range(self):
        assert p("=SUM(A1:A2)") == fn("SUM", rng("Sheet1!A1", "Sheet1!A2"))

    def test_sum_plus_cell(self):
        result = p("=SUM(A1:A9)+Sheet2!C5*1.1")
        assert result == fn(
            "ADD",
            fn("SUM", rng("Sheet1!A1", "Sheet1!A9")),
            fn("MUL", ref("Sheet2!C5"), 1.1),
        )

    def test_zero_arg_function(self):
        assert p("=NOW()") == fn("NOW")

    def test_nested_functions(self):
        assert p("=SUM(ABS(A1),B1)") == fn(
            "SUM",
            fn("ABS", ref("Sheet1!A1")),
            ref("Sheet1!B1"),
        )

    def test_if_three_args(self):
        assert p("=IF(A1>0,B1,C1)") == fn(
            "IF",
            fn("GT", ref("Sheet1!A1"), 0),
            ref("Sheet1!B1"),
            ref("Sheet1!C1"),
        )

    def test_leading_equals_optional(self):
        with_eq = p("=SUM(A1:A2)")
        without_eq = parse_formula("SUM(A1:A2)", SHEET)
        assert with_eq == without_eq
