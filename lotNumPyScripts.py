import xlwings as xw
import psycopg2

def lotNumtoPG(blend_pn, description, lot_number, qty, date_created, line):

    cnxnPG = psycopg2.connect('postgresql://postgres:blend2021@localhost:5432/blendversedb')
    cursPG = cnxnPG.cursor()
    sqlString = "INSERT INTO core_lotnumrecord (part_number, description, lot_number, quantity, date_created, line) VALUES ('" + blend_pn + "', '" + description + "', '" + lot_number + "', '" + qty + "', '" + date_created + "', '" + line + "')"
    cursPG.execute(sqlString)

    cnxnPG.commit()
    cursPG.close()
    cnxnPG.close()

def blendScheduler(blend_pn, description, qty, totes_needed, blend_area, lot_id):
    cnxnPG = psycopg2.connect('postgresql://postgres:blend2021@localhost:5432/blendversedb')
    cursPG = cnxnPG.cursor()
    if blend_area == 'Desk1':
        tblName = 'core_deskoneschedule'
    elif blend_area == 'Desk2':
        tblName = 'core_desktwoschedule'
    cursPG.execute('SELECT MAX("order") FROM ' + tblName)
    if cursPG.fetchone()[0]:
        next_order = str(cursPG.fetchone()[0] + 1)
    else:
        next_order = '0'

    sqlString = "INSERT INTO " + tblName + """ ("order", blend_pn, description, quantity, totes_needed, blend_area, lot_id) VALUES ('"""
    sqlString += next_order + "', '" + blend_pn + "', '" + description + "', '"  + qty + "', '" + totes_needed + "', '" + blend_area + "', (SELECT lot_number from core_lotnumrecord where lot_number='" + lot_id + "') )"
    cursPG.execute(sqlString)

    cnxnPG.commit()
    cursPG.close()
    cnxnPG.close()
