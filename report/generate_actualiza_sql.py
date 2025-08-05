# -*- coding: utf-8 -*-
from tempfile import NamedTemporaryFile
from openpyxl import Workbook
from openpyxl.styles import Color, Fill, Font, Alignment
from openpyxl.cell import Cell
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook
from datetime import datetime
import base64
from openerp import api, models, fields

class GenerateStockWizard(models.Model):

    _name = "sql_process_actualiza.sql_actualiza_execute"
    data = fields.Binary('File', readonly=True)
    name = fields.Char('File Name', readonly=True)
    state = fields.Selection([('choose', 'choose'),
                              ('get', 'get')], default='choose')

    categoria_employee_id = fields.Many2one('hr.employee.category', string='Tipo de Planilla', required=False)
    imprime = fields.Boolean(string='¿Imprime? ', default=True, required=True)
    actualiza = fields.Boolean(string='¿Actualiza? ', default=True, required=True)    
    employee_id = fields.Many2one('hr.employee', string='Empleado', required=False)    
    fecha_desde = fields.Date(string="Fecha Desde", default=fields.Date.today, required=True)
    fecha_hasta = fields.Date(string="Fecha Hasta", default=fields.Date.today, required=True)

    @api.multi
    def generate_file(self):
        
        this = self.id
        
        fileobj = NamedTemporaryFile('w+b')
        xlsfile = fileobj.name
        fileobj.close()
        
        user = self.env.user
        company_id = user.company_id.id
        
#--------------------------------------------------------#
# -------------- Actualiza Empleado ---------------------#
#--------------------------------------------------------#

        # hr_employee.dias_laborados_ano (No tomar en cuenta fines de semana, asuetos o feriados en
        #                                 los que el trabajador haya gozado de descanso).
        # hr_employee.wage (ingresar salario mensual pactado en el contrato sin bonificaciones). Se tima el wage.
        # hr_employee.salario_anual_nominal (Ingresar salario mensual pactado en el contrato sin bonificaciones). Se suman los sueldos ordinarios recibidos en el año.
        # hr_employee.bonificacion_decreto (Si el trabajador recibe la bonificación mensual según decreto colocar Q.250.00. De lo contrario dejar la casialla en blanco.
        # Si el trabajador recibe más de los Q250.00 colocar la diferencia en "Bonificaciones adicionales). Se toma de hr.contract.otra_bonificacion.
        # hr_employee.valor_extra (Ingrear el monto en quetzales que paga la empresa por cada hora extra realizada. Si no realizó horas extras dejar la casilla en blanco).
        # hr_employee.extras_anuales (Ingrese el número total de horas extras que realizadas por el trabajador durante el año. Si no realizó horas extras dejar la casilla en blanco).
        # hr_employee.aguinaldo (Ingrese el monto pagado al trabajador por concepto de aguinaldo si aplica. De lo contrario dejar la casilla en blanco).
        # hr_employee.bono14 (Ingrese el monto pagado en quetzales al trabajador por concepto de bono 14, si aplica. De lo contrario dejar la casilla en blanco).
        # hr_employee.viaticos (Ingrese el monto pagado en quetzales al trabajador por concepto de víaticos, si aplica. De lo contrario dejar la casilla en blanco).
        # hr_employee.bonificaciones (Ingrese el monto pagado en quetzales al trabajador por concepto de bonificaciones adicionales, si aplica. De lo contratio dejar la casilla en blanco).
        # hr_employee.retribucion_por_comision (Ingrese el monto pagado en quetzales al trabajador por concepto de comisiones, si aplica. De lo contratio dejar la casilla en blanco).
        # hr_employee.retribucion_vacaciones (Ingrese el monto pagado en quetzales al trabajador por concepto de vacaciones, si aplica. De lo contratio dejar la casilla en blanco).
        # hr_employee.retribucion_indemnizacion (Ingrese el monto pagado en quetzales al trabajador por concepto de indemnización, si aplica. De lo contratio dejar la casilla en blanco).

        if self.categoria_employee_id.id:
            
            company_id = user.company_id.id
        
        else:
            
            company_id = user.company_id.id
        
        # Actualiza días laborados en empleado.
        if self.actualiza:   
        
            # Actualiza valores anuales en empleado,
            sql = """
                  
                    WITH UltimoSalario AS (
                        SELECT 
                            he.id AS employee_id,
                            hpl.total AS ultimo_salario,
                            ROW_NUMBER() OVER (
                                PARTITION BY he.id 
                                ORDER BY hpr.date_start DESC, hpr.id DESC
                            ) AS rn
                        FROM 
                            hr_payslip_run hpr 
                            JOIN hr_payslip hp ON hp.payslip_run_id = hpr.id
                            JOIN hr_contract hc ON hc.id = hp.contract_id
                            JOIN res_company rc ON rc.id = hp.company_id
                            JOIN hr_employee he ON hp.employee_id = he.id
                            JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                            JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id
                            
                """

            if self.employee_id.id:

                sql += """
                            WHERE 
                                he.id = %s 
                                AND hsr.code IN ('SO', 'SO2')
                                AND hpr.date_start >= %s 
                                AND hpr.date_start <= %s
                                AND hpr.date_end >= hc.date_start 
                                AND (hc.date_end >= %s OR hc.date_end IS NULL)
                        ),
                        CalculatedData AS (
                            SELECT 
                                he.id AS employee_id,
                                SUM(CASE WHEN hsr.code = 'raDLA' THEN hpl.total ELSE 0 END) AS dias,
                                SUM(CASE WHEN hsr.code IN ('ORDQ1', 'ORDQ2') THEN hpl.total ELSE 0 END) AS salario_ordinario_anual,
                                MAX(hc.otra_bonificacion) AS bonificacion_decreto,
                                COALESCE(us.ultimo_salario, 0) AS ultimo_salario
                            FROM 
                                hr_payslip_run hpr 
                                JOIN hr_payslip hp ON hp.payslip_run_id = hpr.id
                                JOIN hr_contract hc ON hc.id = hp.contract_id
                                JOIN res_company rc ON rc.id = hp.company_id
                                JOIN hr_employee he ON hp.employee_id = he.id
                                JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                                JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id
                                LEFT JOIN UltimoSalario us ON us.employee_id = he.id AND us.rn = 1
                            WHERE 
                                he.id = %s 
                                AND hsr.code IN ('raDLA', 'ORDQ1', 'ORDQ2', 'LiqAG', 'LiqB14', 'SO', 'SO2')
                                AND hpr.date_start >= %s 
                                AND hpr.date_start <= %s
                                AND hpr.date_end >= hc.date_start 
                                AND (hc.date_end >= %s OR hc.date_end IS NULL)
                            GROUP BY 
                                he.id, us.ultimo_salario
                        )
                        UPDATE hr_employee em
                        SET 
                            dias_laborados_ano = cd.dias,
                            salario_anual_nominal = cd.salario_ordinario_anual,
                            bonificacion_decreto = cd.bonificacion_decreto,
                            x_ultimo_salario_ordinario = cd.ultimo_salario
                        FROM 
                            CalculatedData cd
                        WHERE 
                            em.id = cd.employee_id;

                    """

                self.env.cr.execute(sql, (self.employee_id.id, self.fecha_desde, self.fecha_hasta, self.fecha_desde, self.employee_id.id, self.fecha_desde, self.fecha_hasta, self.fecha_desde))

            else:

                sql += """
                                WHERE 
                                he.id >= 1 
                                AND hsr.code IN ('SO', 'SO2')
                                AND hpr.date_start >= %s 
                                AND hpr.date_start <= %s
                                AND hpr.date_end >= hc.date_start 
                                AND (hc.date_end >= %s OR hc.date_end IS NULL)
                        ),
                        CalculatedData AS (
                            SELECT 
                                he.id AS employee_id,
                                SUM(CASE WHEN hsr.code = 'raDLA' THEN hpl.total ELSE 0 END) AS dias,
                                SUM(CASE WHEN hsr.code IN ('ORDQ1', 'ORDQ2') THEN hpl.total ELSE 0 END) AS salario_ordinario_anual,
                                MAX(hc.otra_bonificacion) AS bonificacion_decreto,
                                COALESCE(us.ultimo_salario, 0) AS ultimo_salario
                            FROM 
                                hr_payslip_run hpr 
                                JOIN hr_payslip hp ON hp.payslip_run_id = hpr.id
                                JOIN hr_contract hc ON hc.id = hp.contract_id
                                JOIN res_company rc ON rc.id = hp.company_id
                                JOIN hr_employee he ON hp.employee_id = he.id
                                JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                                JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id
                                LEFT JOIN UltimoSalario us ON us.employee_id = he.id AND us.rn = 1
                            WHERE 
                                he.id >= 1 
                                AND hsr.code IN ('raDLA', 'ORDQ1', 'ORDQ2', 'LiqAG', 'LiqB14', 'SO', 'SO2')
                                AND hpr.date_start >= %s 
                                AND hpr.date_start <= %s
                                AND hpr.date_end >= hc.date_start 
                                AND (hc.date_end >= %s OR hc.date_end IS NULL)
                            GROUP BY 
                                he.id, us.ultimo_salario
                        )
                        UPDATE hr_employee em
                        SET 
                            dias_laborados_ano = cd.dias,
                            salario_anual_nominal = cd.salario_ordinario_anual,
                            bonificacion_decreto = cd.bonificacion_decreto,
                            x_ultimo_salario_ordinario = cd.ultimo_salario
                        FROM 
                            CalculatedData cd
                        WHERE 
                            em.id = cd.employee_id;

                          """

                self.env.cr.execute(sql, (self.fecha_desde, self.fecha_hasta, self.fecha_desde, self.fecha_desde, self.fecha_hasta, self.fecha_desde))

#--------------------------------------------------------#
# ---------- Resumen por Empleado -----------------------#
#--------------------------------------------------------#

        if self.imprime:
        
            wb = Workbook()
            
            ws = wb.active
            
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToHeight = 0
            ws.page_setup.fitToWidth = 1
    
            al1 = Alignment(horizontal="left", vertical="center")
            font = Font(size=12, bold=True, italic=True, color="E0D7D7") #, color="4777AD"
            fill = PatternFill("solid", fgColor="4777AD")
            thin = Side(border_style="thin", color="000000")
            double = Side(border_style="double", color="000000")
            
            border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    
            ws.title = "Resumen Empleado"
             
            self.style_range(ws, 'A1:J1', border=border, alignment=al1, font=font, fill=fill)
        
            ws.column_dimensions['A'].width = 40
            ws.column_dimensions['B'].width = 40
            ws.column_dimensions['C'].width = 13
            ws.column_dimensions['D'].width = 14
            ws.column_dimensions['E'].width = 8
            ws.column_dimensions['F'].width = 16
            ws.column_dimensions['G'].width = 16
            ws.column_dimensions['H'].width = 16
            ws.column_dimensions['I'].width = 16
            ws.column_dimensions['J'].width = 16
            
            ws['A1'].value = "Empleado"
            ws['B1'].value = "Departamento"
            ws['C1'].value = "Fecha Contrato"
            ws['D1'].value = "Salario Actual"
            ws['E1'].value = "Días"
            ws['F1'].value = "Salario"
            ws['G1'].value = "Bonificación"
            ws['H1'].value = "Ordinario"
            ws['I1'].value = "Otros Ingresos"
            ws['J1'].value = "Sueldo"

            # daDLA Días Laborados 2da quincena.
            # DiasLM Días laborados mensuales. Solo aparece en las segundas quincenas.
            # raDLA días laborados primera quincena. Aparece en cada quincena.
            # ORDQ1 salario ordinario primera quincena.
            # ORDQ2 salario ordinario segunda quincena.
            # LiqAG liquido a recibir aguinaldo.
            # LiqB14 liquido a recibir bono 14.
            # SO salario.
            # SO2 salario base.
            # BO Bonificación decreto Q250.00 372001 primera quincena.
            # daBO Bonificación decreto Q250.00 372001 2da quincena.
            # HextQ1 Cantidad horas extras primera quincena.
            # HextQ2 Cantidad horas extras segunda quincena.
            # Valor hora extra round((contract.wage / 30 / 8 * 1.5), 2), tomar en cuenta el último salario ordinario recibido.
            # LiqAG Líquido a recibir aguinaldo.
            # LiqB14 Líquido a recibir bono 14.
            # No pagan comisiones (debe agregarse regla salarial si se necesita).
            # VTCQ1 Víaticos 1ra. quincena.
            # VTCQ2 Víaticos 2da. quincena.
            #if employee.x_aplica_bono_fijo:
            #    result = (((contract.ingresos1 + contract.ingresos4) * employee.dias_laborados_ano) / 360) / 12
            #if employee.x_aplica_bono_variable:
            #    result = (((contract.ingresos2 + contract.ingresos3) * employee.dias_laborados_ano) / 360) / 12
            # ValorvacQ1 Valor vacaciones primera quincena.
            # ValorvacQ2 Valor vacaciones segunda quincena.
            # No tenemos la retribución por indemnización, se debe crear una esctructura salarial.

            # Detalle la actualización realizada a cada empleado.
            sql = """        
                    SELECT r.contract_id, he.id AS employee_id, he.name AS nombre_empleado, he.department_id
                         , CONCAT(COALESCE(hda.name,''), (CASE WHEN hda.name isnull THEN '' ELSE '/' END), COALESCE(hdp.name,''), (CASE WHEN hdp.name isnull THEN '' ELSE '/' END), COALESCE(hd.name,'')) AS departamento
                         , he.bonificacion_decreto, he.bonificaciones, hc.date_start AS fecha_inicio_contrato, hc.ingresos1, hc.ingresos4, r.dias, hc.wage AS salario_actual
                         , r.bonificacion_acumulada, r.sueldo_acumulado, r.salario_ordinario_anual, hc.ingresos1 + hc.ingresos4 AS bonificacion_adicional
                      FROM (SELECT d.contract_id, SUM(d.dias) AS dias, (SUM(d.ordinario_quincena1)+SUM(d.ordinario_quincena2)) AS salario_ordinario_anual, SUM(bonificacion) AS bonificacion_acumulada 
                                 , SUM(bonificacion) AS sueldo_acumulado
                           FROM (SELECT hp.contract_id,                     
                            CASE WHEN hsr.code = 'raDLA' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS dias, 
                            CASE WHEN hsr.code = 'SO' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS salario_mensual,                           
                            CASE WHEN hsr.code = 'ORDQ1' THEN SUM(hpl.total)
                                ELSE 0.00
                            END AS ordinario_quincena1,
                            CASE WHEN hsr.code = 'ORDQ2' THEN SUM(hpl.total)
                                ELSE 0.00
                            END AS ordinario_quincena2,
                            CASE WHEN hsr.code = 'LiqAG' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS aguinaldo,
                            CASE WHEN hsr.code = 'SO' THEN SUM(hpl.total)
                            ELSE 0.00
                            END AS bonificacion, 
							CASE WHEN hsr.code = 'LiqB14' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS bono_14
                     FROM hr_payslip_run hpr 
                     JOIN hr_payslip hp ON hp.payslip_run_id = hpr.id
                     JOIN hr_contract hc ON hp.contract_id = hc.id
                     JOIN res_company rc ON hp.company_id = rc.id                  
                     JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                     JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id    
                    """
                
            if self.employee_id.id:
            
                sql += """               
                   
                     WHERE rc.id = %s
                       AND hp.employee_id = %s AND hsr.code IN ('raDLA', 'SO', 'ORDQ1', 'ORDQ2', 'LiqAG', 'LiqB14')
                       AND hpr.date_start >= %s AND hpr.date_start <= %s
                       AND hpr.date_end >= hc.date_start AND hc.date_end isnull
                   GROUP BY hp.contract_id, hsr.code) d
                   GROUP BY d.contract_id) r
                   JOIN hr_contract hc ON r.contract_id = hc.id
                   LEFT JOIN hr_employee he ON he.id = hc.employee_id
                   LEFT JOIN hr_department hd ON he.department_id = hd.id
                   LEFT JOIN hr_department hdp ON hd.parent_id = hdp.id
                   LEFT JOIN hr_department hda ON hdp.parent_id = hda.id
                   ORDER BY he.name;        
                """        
        
                self.env.cr.execute(sql, (company_id, self.employee_id.id, self.fecha_desde, self.fecha_hasta,))
                                     
            else:
            
                sql += """
                
                    WHERE rc.id = %s
                       AND hp.employee_id >= 1 AND hsr.code IN ('raDLA', 'SO', 'ORDQ1', 'ORDQ2', 'LiqAG', 'LiqB14')
                       AND hpr.date_start >= %s AND hpr.date_start <= %s
                       AND hpr.date_end >= hc.date_start AND hc.date_end isnull
                   GROUP BY hp.contract_id, hsr.code) d
                   GROUP BY d.contract_id) r
                   JOIN hr_contract hc ON r.contract_id = hc.id
                   LEFT JOIN hr_employee he ON he.id = hc.employee_id
                   LEFT JOIN hr_department hd ON he.department_id = hd.id
                   LEFT JOIN hr_department hdp ON hd.parent_id = hdp.id
                   LEFT JOIN hr_department hda ON hdp.parent_id = hda.id
                   ORDER BY he.name;        
                        
                """        
                
                self.env.cr.execute(sql, (company_id, self.fecha_desde, self.fecha_hasta,))
            
            row = 2
            
            # Inicializa total.        
            wtotal_salario = 0
            wtotal_bonificacion = 0
            wtotal_ordinario = 0
            wtotal_otros_ingresos = 0
            wtotal_sueldo = 0
            
            for query_line in self.env.cr.dictfetchall():
              
                ws.cell(row=row, column=1).value  = query_line['nombre_empleado'] or None;
                
                ws.cell(row=row, column=2).value  = query_line['departamento'] or None;

                fs = query_line['fecha_inicio_contrato'].strftime('%Y-%m-%d').split('-')
                _fecha = ((fs[2]) + "/" + fs[1] + "/" + fs[0])

                ws.cell(row=row, column=3).value  = _fecha or None;
                
                ws.cell(row=row, column=4).value  = round(query_line['salario_actual'],0) or 0;
                ws.cell(row=row, column=4).number_format = '#,##0.00'            
                
                ws.cell(row=row, column=5).value  = round(query_line['dias'],0) or 0;
                ws.cell(row=row, column=5).number_format = '#,##0'
        
                ws.cell(row=row, column=6).value  = round(query_line['salario_ordinario_anual'],2) or 0;
                ws.cell(row=row, column=6).number_format = '#,##0.00'
                
                ws.cell(row=row, column=7).value  = round(query_line['bonificacion_adicional'],2) or 0;
                #ws.cell(row=row, column=7).value = 0;
                ws.cell(row=row, column=7).number_format = '#,##0.00'
                            
                ws.cell(row=row, column=8).value  = round(query_line['bonificacion_decreto'],2) or 0;
                ws.cell(row=row, column=8).number_format = '#,##0.00'
                
                ws.cell(row=row, column=9).value  = round(query_line['bonificacion_acumulada'],2) or 0;
                ws.cell(row=row, column=9).number_format = '#,##0.00'
                
                ws.cell(row=row, column=10).value  = round(query_line['sueldo_acumulado'],2) or 0;
                ws.cell(row=row, column=10).number_format = '#,##0.00'
                
                wtotal_salario          += query_line['salario_ordinario_anual']
                wtotal_bonificacion     += query_line['bonificacion_adicional']
                wtotal_ordinario        += query_line['bonificacion_decreto']
                wtotal_otros_ingresos   += query_line['bonificacion_acumulada']
                wtotal_sueldo           += query_line['sueldo_acumulado']
                
                row += 1
                  
            # Totales del Reporte.
            ws.cell(row=row, column=1).value  = 'Total: '
            ws.cell(row=row, column=1).font = Font(size = 13, bold = True)  
        
            ws.cell(row=row, column=6).value  = wtotal_salario
            ws.cell(row=row, column=6).number_format='#,##0.00'
            ws.cell(row=row, column=6).font = Font(size = 13, bold = True) 
            
            ws.cell(row=row, column=7).value  = wtotal_bonificacion
            ws.cell(row=row, column=7).number_format='#,##0.00'
            ws.cell(row=row, column=7).font = Font(size = 13, bold = True) 
            
            ws.cell(row=row, column=8).value  = wtotal_ordinario
            ws.cell(row=row, column=8).number_format='#,##0.00'
            ws.cell(row=row, column=8).font = Font(size = 13, bold = True) 
        
            ws.cell(row=row, column=9).value  = wtotal_otros_ingresos
            ws.cell(row=row, column=9).number_format='#,##0.00'
            ws.cell(row=row, column=9).font = Font(size = 13, bold = True)
                            
            ws.cell(row=row, column=10).value  = wtotal_sueldo
            ws.cell(row=row, column=10).number_format='#,##0.00'
            ws.cell(row=row, column=10).font = Font(size = 13, bold = True)
                
#--------------------------------------------------------#
# ---------- Detalle por Empleado -----------------------#
#--------------------------------------------------------#

            ws = wb.create_sheet()
            
            ws.title = "Detalle Empleado"
            
            self.style_range(ws, 'A1:l1', border=border, alignment=al1, font=font, fill=fill)
    
            ws.column_dimensions['A'].width = 40
            ws.column_dimensions['B'].width = 44
            ws.column_dimensions['C'].width = 40
            ws.column_dimensions['D'].width = 13
            ws.column_dimensions['E'].width = 13
            ws.column_dimensions['F'].width = 13
            ws.column_dimensions['G'].width = 8
            ws.column_dimensions['H'].width = 16
            ws.column_dimensions['I'].width = 16
            ws.column_dimensions['J'].width = 16
            ws.column_dimensions['K'].width = 16
            ws.column_dimensions['l'].width = 16  
                         
            ws['A1'].value = "Empleado"
            ws['B1'].value = "Nómina"
            ws['C1'].value = "Departamento"
            ws['D1'].value = "Inicio Contrato"
            ws['E1'].value = "Inicio"
            ws['F1'].value = "Final"
            ws['G1'].value = "Días"
            ws['H1'].value = "Salario"
            ws['I1'].value = "Bonificación"
            ws['J1'].value = "Ordinario"
            ws['K1'].value = "Otros Ingresos"
            ws['L1'].value = "Sueldo"
            
            # Detalla la actualización realizada a cada empleado.  
            sql = """           

                SELECT r.contract_id, hc.date_start AS fecha_inicio_contrato, hc.wage AS salario_actual, he.id AS employee_id, he.name AS nombre_empleado
                     , he.department_id, CONCAT(COALESCE(hda.name,''), (CASE WHEN hda.name isnull THEN '' ELSE '/' END), COALESCE(hdp.name,''), (CASE WHEN hdp.name isnull THEN '' ELSE '/' END), COALESCE(hd.name,'')) AS departamento
                     , r.id, hpr.name AS nombre_nomina, hpr.date_start AS fecha_inicio_nomina, hpr.date_end AS fecha_final_nomina
                     , hc.ingresos1, hc.ingresos4, r.dias, hc.wage AS salario_actual, he.bonificacion_decreto, he.bonificaciones
                     , r.bonificacion_acumulada, r.sueldo_acumulado, r.salario_ordinario_anual, hc.ingresos1 + hc.ingresos4 AS bonificacion_adicional
                  FROM (SELECT d.contract_id, d.id,
                        SUM(d.dias) AS dias, (SUM(d.ordinario_quincena1)+SUM(d.ordinario_quincena2)) AS salario_ordinario_anual, SUM(d.bonificacion) AS bonificacion_acumulada 
                                 , SUM(d.bonificacion) AS sueldo_acumulado    
                     FROM (SELECT hp.contract_id, hpr.id,
                            CASE WHEN hsr.code = 'raDLA' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS dias, 
                            CASE WHEN hsr.code = 'SO' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS salario_mensual,                           
                            CASE WHEN hsr.code = 'ORDQ1' THEN SUM(hpl.total)
                                ELSE 0.00
                            END AS ordinario_quincena1,
                            CASE WHEN hsr.code = 'ORDQ1' THEN SUM(hpl.total)
                                ELSE 0.00
                            END AS ordinario_quincena2,
                            CASE WHEN hsr.code = 'LiqAG' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS aguinaldo, 
							CASE WHEN hsr.code = 'LiqB14' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS bono_14,
                            CASE WHEN hsr.code = 'LiqB14' THEN SUM(hpl.total)
                                ELSE 0.0
                            END AS bonificacion
                             FROM hr_payslip_run hpr 
                             JOIN hr_payslip hp ON hpr.id = hp.payslip_run_id
                             JOIN hr_contract hc ON hp.contract_id = hc.id
                             JOIN res_company rc ON hp.company_id = rc.id
                             JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                             JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id                   
                      
            """ 
                
            if self.employee_id:
            
                sql += """   
                 
                      WHERE rc.id = %s
                          AND hp.employee_id = %s (hsr.code = 'raDLA' OR hsr.code = 'SO' OR hsr.code = 'ORDQ1' OR hsr.code = 'ORDQ2'
                               OR hsr.code = 'LiqAG' OR hsr.code = 'LiqB14')
                          AND hpr.date_start >= %s AND hpr.date_start <= %s
                          AND hpr.date_end >= hc.date_start AND hc.date_end isnull
                     GROUP BY hp.contract_id, hpr.id, hsr.name) d                       
                     GROUP BY d.contract_id, d.id) r
                 JOIN hr_contract hc ON r.contract_id = hc.id
                     JOIN hr_payslip_run hpr ON r.id = hpr.id 
                 LEFT JOIN hr_employee he ON he.id = hc.employee_id
                 LEFT JOIN hr_department hd ON he.department_id = hd.id
                 LEFT JOIN hr_department hdp ON hd.parent_id = hdp.id
                 LEFT JOIN hr_department hda ON hdp.parent_id = hda.id
                 ORDER BY r.contract_id, r.id;
                      
                """        
    
                self.env.cr.execute(sql, (company_id, self.employee_id.id, self.fecha_desde, self.fecha_hasta,))
                                     
            else:
            
                sql += """
                        
                        WHERE rc.id = %s
                          AND hp.employee_id >= 1 AND (hsr.code = 'raDLA' OR hsr.code = 'SO' OR hsr.code = 'ORDQ1' OR hsr.code = 'ORDQ2'
                               OR hsr.code = 'LiqAG' OR hsr.code = 'LiqB14')
                          AND hpr.date_start >= %s AND hpr.date_start <= %s
                          AND hpr.date_end >= hc.date_start AND hc.date_end isnull
                     GROUP BY hp.contract_id, hpr.id, hsr.code) d                       
                     GROUP BY d.contract_id, d.id) r
                 JOIN hr_contract hc ON r.contract_id = hc.id
                     JOIN hr_payslip_run hpr ON r.id = hpr.id 
                   LEFT JOIN hr_employee he ON he.id = hc.employee_id
                   LEFT JOIN hr_department hd ON he.department_id = hd.id
                   LEFT JOIN hr_department hdp ON hd.parent_id = hdp.id
                   LEFT JOIN hr_department hda ON hdp.parent_id = hda.id
                 ORDER BY r.contract_id, r.id;
                                
                """        
                
                self.env.cr.execute(sql, (company_id, self.fecha_desde, self.fecha_hasta,))
            
            row = 2
            
            # Inicializa total.        
            wtotal_salario = 0
            wtotal_bonificacion = 0
            wtotal_ordinario = 0
            wtotal_sueldo = 0
            
            for query_line in self.env.cr.dictfetchall():
    
                ws.cell(row=row, column=1).value  = query_line['nombre_empleado'] or None;
                            
                ws.cell(row=row, column=2).value  = query_line['nombre_nomina'] or None;
    
                ws.cell(row=row, column=3).value  = query_line['departamento'] or None;

                fs = query_line['fecha_inicio_contrato'].strftime('%Y-%m-%d').split('-')
                _fecha = ((fs[2]) + "/" + fs[1] + "/" + fs[0])
                
                ws.cell(row=row, column=4).value  = _fecha or None;

                fs = query_line['fecha_inicio_nomina'].strftime('%Y-%m-%d').split('-')
                _fecha = ((fs[2]) + "/" + fs[1] + "/" + fs[0])

                ws.cell(row=row, column=5).value  = _fecha or None;

                fs = query_line['fecha_final_nomina'].strftime('%Y-%m-%d').split('-')
                _fecha = ((fs[2]) + "/" + fs[1] + "/" + fs[0])

                ws.cell(row=row, column=6).value  = _fecha or None;
                
                ws.cell(row=row, column=7).value  = round(query_line['dias'],0) or 0;
                ws.cell(row=row, column=7).number_format = '#,##0'
    
                ws.cell(row=row, column=8).value  = round(query_line['salario_ordinario_anual'],2) or 0;
                ws.cell(row=row, column=8).number_format = '#,##0.00'

                ws.cell(row=row, column=9).value = round(query_line['bonificacion_adicional'], 2) or 0;
                ws.cell(row=row, column=9).number_format = '#,##0.00'
                
                ws.cell(row=row, column=10).value  = round(query_line['bonificacion_decreto'],2) or 0;
                ws.cell(row=row, column=10).number_format = '#,##0.00'
    
                ws.cell(row=row, column=11).value  = round(query_line['bonificacion_acumulada'],2) or 0;
                ws.cell(row=row, column=11).number_format = '#,##0.00'
                            
                ws.cell(row=row, column=12).value  = round(query_line['salario_ordinario_anual'],2) or 0;
                ws.cell(row=row, column=12).number_format = '#,##0.00'
                
                wtotal_salario          += query_line['salario_ordinario_anual']
                wtotal_bonificacion     += query_line['bonificacion_adicional']
                wtotal_ordinario        += query_line['bonificacion_acumulada']
                wtotal_otros_ingresos   += query_line['bonificacion_acumulada']
                wtotal_sueldo           += query_line['salario_ordinario_anual']
                
                row += 1
                  
            # Totales del Reporte.
            ws.cell(row=row, column=1).value  = 'Total: '
            ws.cell(row=row, column=1).font = Font(size = 13, bold = True)  
    
            ws.cell(row=row, column=8).value  = wtotal_salario
            ws.cell(row=row, column=8).number_format='#,##0.00'
            ws.cell(row=row, column=8).font = Font(size = 13, bold = True) 
            
            ws.cell(row=row, column=9).value  = wtotal_bonificacion
            ws.cell(row=row, column=9).number_format='#,##0.00'
            ws.cell(row=row, column=9).font = Font(size = 13, bold = True) 
            
            ws.cell(row=row, column=10).value  = wtotal_ordinario
            ws.cell(row=row, column=10).number_format='#,##0.00'
            ws.cell(row=row, column=10).font = Font(size = 13, bold = True) 
    
            ws.cell(row=row, column=11).value  = wtotal_otros_ingresos
            ws.cell(row=row, column=11).number_format='#,##0.00'
            ws.cell(row=row, column=11).font = Font(size = 13, bold = True) 
                            
            ws.cell(row=row, column=12).value  = wtotal_sueldo
            ws.cell(row=row, column=12).number_format='#,##0.00'
            ws.cell(row=row, column=12).font = Font(size = 13, bold = True) 
            
            wb.save(filename=xlsfile)
    
            spreadsheet_file = open(xlsfile, "rb")
            binary_data = spreadsheet_file.read()
            spreadsheet_file.close()
            out = base64.b64encode(binary_data)
    
            self.write({
                'state': 'get',
                'name': "actualiza_empleado_spreadsheet.xlsx",
                'data': out
            })
    
            return {
                'type': 'ir.actions.act_window',
                'res_model': 'sql_process_actualiza.sql_actualiza_execute',
                'view_mode': 'form',
                'view_type': 'form',
                'res_id': this,
                'views': [(False, 'form')],
                'target': 'new',
            }

    def style_range(self, ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
        """
        Apply styles to a range of cells as if they were a single cell.
        :param ws:  Excel worksheet instance
        :param range: An excel range to style (e.g. A1:F20)
        :param border: An openpyxl Border
        :param fill: An openpyxl PatternFill or GradientFill
        :param font: An openpyxl Font object
        """

        top = Border(top=border.top)
        left = Border(left=border.left)
        right = Border(right=border.right)
        bottom = Border(bottom=border.bottom)

        first_cell = ws[cell_range.split(":")[0]]
        rows = list(ws[cell_range])
        for cell in rows[0]:
            cell.border = top
        for cell in rows[-1]:
            cell.border = bottom
        for row in rows:
            l = row[0]
            r = row[-1]
            l.border =  left
            r.border =  right
            if fill:
                for c in row:
                    c.fill = fill
            if font:
                for c in row:
                    c.font = font

            if alignment:
                for c in row:
                    c.alignment = alignment
        