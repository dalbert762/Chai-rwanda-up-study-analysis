"""Path setup for the analytics scripts, which live one folder down.

Every script in this folder starts with:

    import _bootstrap as B

That does three things:

  * puts the project root on sys.path, so `import analyze_medicine_assessment`
    still works;
  * puts THIS folder on sys.path, so the scripts can import each other
    (Completion_by_geography imports Site_coverage_check);
  * changes the working directory to the project root, because the analysis
    modules open the merged exports by relative filename.

Outputs still belong here, not in the root, so each script writes to
`B.out("Name.xlsx")` rather than to a bare filename.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

for p in (ROOT, HERE):
    if p not in sys.path:
        sys.path.insert(0, p)

# the analysis modules read Medicine_stock_assessment_export.xlsx and friends
# by relative name, so the scripts have to run from the root
os.chdir(ROOT)


def out(name):
    """Absolute path for an output file, inside the analytics folder."""
    return os.path.join(HERE, name)
