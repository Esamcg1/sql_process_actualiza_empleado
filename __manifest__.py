# -*- coding: utf-8 -*-
# © <2015> <Miguel Chuga>
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).
{
    "name": "Actualiza Empleado",
    "summary": "Actualiza información de nómina en empleado.",
    "version": "8.0.1.0.0",
    "category": "Report",
    "website": "https://mcsistemas.net",
    "author": "Miguel Chuga,"
              "MC-Sistemas",
    "license": "AGPL-3",
    "application": False,
    "installable": True,
    "depends": [
        "base",
        "account",
        'hr',
        "mc_nomina_gt",
    ],
    "data": [
        'report/generate_actualiza_sql.xml',
    ],
    "demo": [
    ],
    'installable': True,
    'auto_install': False,
}