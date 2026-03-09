{
    "name": "臺灣營業稅申報模組",
    "version": "1.5.0",
    "category": "Accounting/Localizations",
    "summary": "臺灣營業稅申報與匯出",
    "license": "LGPL-3",
    "author": "pAq Computer Enterprise",
    "depends": ["account"],
    "data": [
        "security/ir.model.access.csv",
        "views/res_company_views.xml",
        "views/account_move_views.xml",
        "wizard/vat_filing_wizard_views.xml",
    ],
    "installable": True,
    "application": True,
}
