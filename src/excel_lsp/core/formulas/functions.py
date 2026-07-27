"""Frozen Excel function-name metadata used by formula analysis."""

from __future__ import annotations

from openpyxl.utils import FORMULAE as _OPENPYXL_FORMULAE

# ``openpyxl==3.1.5`` exposes a useful but legacy-biased function list.  Keep
# that import load-bearing, then freeze the names absent from it in Microsoft's
# alphabetical Excel-function catalog as captured on 2026-07-16.  FIELDVALUE,
# SINGLE, ANCHORARRAY, and PY are persisted compatibility names not present in
# that public catalog.  This set is deliberately committed data: formula
# indexing never depends on the network or on the user's installed Excel build.
# Source: Microsoft Support, "Excel functions (alphabetical)",
# article b3944572-255d-4efb-bb96-c6d90033e188.
MODERN_FUNCTIONS = frozenset(
    """
    ACOT ACOTH AGGREGATE ANCHORARRAY ARABIC ARRAYTOTEXT
    BASE BETA.DIST BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND
    BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW
    CALL CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV
    CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT
    CONFIDENCE.NORM CONFIDENCE.T COPILOT COT COTH COVARIANCE.P
    COVARIANCE.S CSC CSCH DAYS DBCS DECIMAL
    DETECTLANGUAGE DROP ENCODEURL ERF.PRECISE ERFC.PRECISE EUROCONVERT
    EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT
    F.TEST FIELDVALUE FILTER FILTERXML FLOOR.MATH FLOOR.PRECISE
    FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY
    FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT
    GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY
    HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH
    IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH
    IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LAMBDA LET
    LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS
    MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NORM.DIST NORM.INV
    NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC
    PERCENTILE.INC PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC
    PERMUTATIONA PHI PIVOTBY POISSON.DIST PY QUARTILE.EXC
    QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ REDUCE
    REGEXEXTRACT REGEXREPLACE REGEXTEST REGISTER.ID RRI SCAN
    SEC SECH SEQUENCE SHEET SHEETS SINGLE
    SKEW.P SORT SORTBY STDEV.P STDEV.S STOCKHISTORY
    SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T
    T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT
    TOCOL TOROW TRANSLATE TRIMRANGE UNICHAR UNICODE
    UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE
    WEIBULL.DIST WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR
    Z.TEST
    """.split()  # noqa: SIM905 - compact frozen data is clearer than 171 quoted lines.
)

BUILTIN_FUNCTIONS = frozenset(name.upper() for name in _OPENPYXL_FORMULAE) | MODERN_FUNCTIONS

VOLATILE_FUNCTIONS = frozenset(
    {
        "CELL",
        "INDIRECT",
        "INFO",
        "NOW",
        "OFFSET",
        "RAND",
        "RANDBETWEEN",
        "RANDARRAY",
        "TODAY",
    }
)

ALWAYS_DYNAMIC_REFERENCE_FUNCTIONS = frozenset({"INDIRECT", "OFFSET"})
CONTEXTUAL_DYNAMIC_REFERENCE_FUNCTIONS = frozenset({"CHOOSE", "INDEX"})


def function_identifier(value: str) -> str:
    """Return the lexical identifier from an open-function token.

    openpyxl sometimes absorbs a range operator and its left endpoint into the
    following function token (for example ``A1:INDEX(``).  Only the suffix is
    the callable identifier.  A leading ``@`` is Excel's implicit-intersection
    operator rather than part of the name.
    """
    identifier = value.strip()
    if identifier.endswith("("):
        identifier = identifier[:-1]
    if ":" in identifier:
        identifier = identifier.rsplit(":", 1)[1]
    return identifier.lstrip("@").strip()


def compatibility_function_identifier(value: str) -> str:
    """Strip Excel compatibility prefixes from one callable identifier."""
    identifier = function_identifier(value)
    # Excel normally stores worksheet-scoped functions as
    # ``_xlfn._xlws.NAME``.  The loop also accepts the defensive reverse or a
    # duplicated compatibility prefix without changing persisted formula text.
    changed = True
    while changed:
        changed = False
        folded = identifier.casefold()
        for prefix in ("_xlfn.", "_xlws."):
            if folded.startswith(prefix):
                identifier = identifier[len(prefix) :]
                changed = True
                break
    return identifier


def normalize_function_name(value: str) -> str:
    """Normalize one stored/call-token function name for matching and display."""
    return compatibility_function_identifier(value).upper()


__all__ = [
    "ALWAYS_DYNAMIC_REFERENCE_FUNCTIONS",
    "BUILTIN_FUNCTIONS",
    "CONTEXTUAL_DYNAMIC_REFERENCE_FUNCTIONS",
    "MODERN_FUNCTIONS",
    "VOLATILE_FUNCTIONS",
    "compatibility_function_identifier",
    "function_identifier",
    "normalize_function_name",
]
