import xlwings as xw
import psycopg2

def lotNumtoPG(blend_pn, description, lot_number, qty, date_created, line):

    cnxnPG = psycopg2.connect('postgresql://postgres:blend2021@localhost:5432/blendversedb')
    cursPG = cnxnPG.cursor()
    sqlString = "INSERT INTO core_lotnumrecord (part_number, description, lot_number, quantity, date_created, line) VALUES ('"+ blend_pn + "', '" + description + "', '" + lot_number + "', '" + qty + "', '" + date_created + "', '" + line + "')"
    cursPG.execute(sqlString)

    cnxnPG.commit()
    cursPG.close()
    cnxnPG.close()

     