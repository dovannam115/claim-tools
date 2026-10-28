# Paths and engine setup shared by server.py (the window) and worker.py (one run of a tool).
import os, sys

APP     = os.path.dirname(os.path.abspath(__file__))
ROOT    = os.path.dirname(APP)
ENGINES = os.path.join(APP, "engines")
DATA    = os.path.join(APP, "data")           # rate cards, rules, Word templates, input templates
DUMP    = os.path.join(APP, "dump")           # token, log, Fabric sign-in record
OUTPUT  = os.path.join(ROOT, "output")        # one folder per run: output\<Tool>\<date time>\

RATE_SOURCE = os.path.join(DATA, "rate_source.xlsm")
RULES       = os.path.join(DATA, "reconcile_tgdd_rules.xlsx")
GCN_TPL     = os.path.join(DATA, "gcnbh_templates")
PIAS_TPL    = os.path.join(DATA, "Pias_Template.xlsx")
TEMPLATES   = {"claim": os.path.join(DATA, "Claim_Input_Template.xlsx"),
               "invoice": os.path.join(DATA, "Invoice_Input_Template.xlsx"),
               "rules": RULES}
AUTH        = os.path.join(DUMP, "fabric_auth_record.json")
ODBC        = "ODBC Driver 18 for SQL Server"
VENDOR      = os.path.join(APP, "vendor")      # mssql-python: SQL Server driver bundled, nothing to install
VENDOR_OK   = os.path.join(VENDOR, "vendor.ok")
SCOPE       = "https://database.windows.net/.default"

if ENGINES not in sys.path: sys.path.insert(0, ENGINES)
if os.path.isdir(VENDOR) and VENDOR not in sys.path: sys.path.insert(1, VENDOR)

# folder name of each tool under output\
TOOLS = {"map": "Map_Claim", "salvage": "Check_Salvage", "invoice": "Check_Invoice",
         "gcngen": "Generate_GCNBH", "gcnsearch": "Search_GCNBH", "gcnprint": "Search_GCNBH",
         "searchdb": "Search_Database", "tgdd": "Reconcile_TGDD", "signin": "_signin",
         "odbc": "_odbc", "update": "_update"}


def core():
    """map_claim_ri set up for the app: errors raise MapError, no PowerShell dialogs, and the
    Fabric sign-in record lives in app\\dump (the engine default is next to the script, which
    an update could overwrite)."""
    import map_claim_ri as c
    c.HEADLESS = True
    c.GUI = False
    c.AUTH = AUTH
    return c


def fabric_cred(c=None):
    """Silent credential from the saved sign-in record, or None when the user must sign in."""
    c = c or core()
    if not os.path.exists(AUTH): return None, ""
    from azure.identity import DeviceCodeCredential, TokenCachePersistenceOptions, AuthenticationRecord
    rec = AuthenticationRecord.deserialize(open(AUTH, encoding="utf-8").read())
    cred = DeviceCodeCredential(client_id=c.CLIENT,
                                cache_persistence_options=TokenCachePersistenceOptions(name="claude_fabric_equip"),
                                authentication_record=rec, disable_automatic_authentication=True)
    return cred, rec.username


# Where "Update" looks for new versions. app\update.json wins when it names a repo; copies handed out
# before 1.0.1 have an empty update.json, so the built-in value is what lets them update at all.
REPO = {"repo": "dovannam115/claim-tools", "branch": "main", "token": ""}


def update_cfg():
    import json
    out = dict(REPO)
    try:
        j = json.load(open(os.path.join(APP, "update.json"), encoding="utf-8"))
        if j.get("repo"): out.update({k: v for k, v in j.items() if v or k == "token"})
    except Exception:
        pass
    return out


def http_get(url, token="", timeout=25):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "ClaimTools", "Cache-Control": "no-cache"})
    if token: req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=timeout) as r: return r.read()


def driver():
    """Which SQL Server driver this copy can use, '' when none. Looks at files only: the server
    process must not load the vendor DLLs, or an update could not replace them."""
    if os.path.exists(VENDOR_OK): return "mssql-python"
    try:
        import pyodbc
        return ODBC if ODBC in pyodbc.drivers() else ""
    except Exception:
        return ""


def connect(c, cred):
    """Fabric connection with the user's own token. mssql-python carries its own ODBC driver, so
    nothing has to be installed on the PC; the system ODBC Driver 18 (engine's connect) is the fallback."""
    import struct
    if os.path.exists(VENDOR_OK):
        import mssql_python
        t = cred.get_token(SCOPE).token.encode("utf-16-le")
        st = struct.pack("<I%ds" % len(t), len(t), t)
        cs = ("SERVER=%s;DATABASE=%s;Encrypt=yes;TrustServerCertificate=no;"
              % (c.SERVER, c.DBNAME))
        return mssql_python.connect(cs, attrs_before={1256: st}, autocommit=True)
    return c.connect()
