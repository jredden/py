import xlrd3
import pymysql

"""
program is used to parse a spread sheet of new models and generate a series of
mysql load files.  

MySqlClient reads the highest ID in the Model table, and generates sequential
unique ideas numbered after the that ID.  This is the only database access
in this program

MySqlClient also hold the attributes to generate SQL load files to later be
inserted into the database.  Yes, they could be inserted directly, this way
there is an audit step where any glitches can be removed prior to loading.

Each column has a class that that is referenced in do spreadsheet.  The classes with 
zero in the name are for the first column only.  Often these are used to initialize 
the next set of sequences.

Do one spread sheet operates on a single spread sheet with multiple tabs.

self.rows is a map that is setup, but was never used.  It is left in for flexibility

"""


class MySqlClient:
    maxModelID = 0
    modelTempID = 0
    model_name = ""
    model_number = ""
    model_short_number = ""
    model_list_order = 0
    model_year = "'2013'"

    temp_id = 0 
    spec_name_english = ""
    spec_desc_english = ""
    spec_list_order = 0
    spread_sheet_id = 0
    dbfile = ""
    dbDetail_file = ""
    dbModelMap_file = ""
    oneTimePerSheet = True
    oneTimeOnly = False

    def __init__(self):
        # --user=jredden --password=nCXL3O2GVAwLBJnT
        self.conn = pymysql.connect(host='localhost', unix_socket='/var/run/mysqld/mysqld.sock', user='jredden', passwd='nCXL3O2GVAwLBJnT', db='suzuki_moto')
        self.cur = self.conn.cursor()
        self.dbfile = open("/tmp/model.sql", 'w')
        self.dbDetail_file = open("/tmp/detail.sql", 'w')
        self.dbModelMap_file = open("/tmp/mpm.sql", 'w')
        self.maxModelID = 0

    def cleanClose(self):
        self.cur.close()
        self.conn.close()
        self.dbfile.close()
        self.dbDetail_file.close()
        self.dbModelMap_file.close()

    def queryMaxModelID(self):
        cur = self.conn.cursor()
        cur.execute("SELECT MAX(ID) FROM Model")
        self.maxModelID = cur.fetchone()[0]
        #print('MaxModelID:', self.maxModelID)

    def getMaxModelID(self):
        return self.maxModelID

    def incrMaxModelID(self):
        self.maxModelID += 1 
        
    def setModelTempID(self, tempID):
        self.modelTempID = tempID

    def getModelTempID(self):
        return self.modelTempID

    def setModelName(self, model_name):
        self.model_name = model_name

    def setModelNumber(self, model_number):
        self.model_number = model_number

    def setModelShortNumber(self, model_short_number):
        self.model_short_number = model_short_number

    def setModelListOrder(self, model_list_order):
        self.model_list_order = model_list_order
    def getModelListOrder(self):
        return self.model_list_order

    def setTempId(self, id):
        self.temp_id = id

    def getTempId(self):
        return self.temp_id
    def incrTempId(self):
        self.temp_id += 1

    def setNameEnglish(self, name_english):
        self.spec_name_english = name_english

    def setDescEnglish(self, description):
        self.spec_desc_english = description

    def setSpecOrder(self, order):
        self.spec_list_order = order

    def setSpreadSheetId(self, id):
        self.spread_sheet_id = id
    def getSpreadSheetId(self):
        return self.spread_sheet_id

    def emitModelSQL(self):
        if(self.oneTimeOnly == False):
           self.oneTimeOnly = True
           return
        sql_insert = "INSERT INTO Model(ID, Name, Number, ShortNumber, Year, ListOrder, MSRP, suzuki_id, msrpsep) VALUES ({}, '{}', '{}', '{}', {}, {}, {}, '{}', {});".format(self.maxModelID, self.model_name, self.model_number, self.model_short_number, self.model_year, self.model_list_order, 0.0, self.maxModelID, 0.0)
        
        print(sql_insert, file=self.dbfile)

    def emitDetailSQL(self):
        sql_insert= "INSERT INTO ModelTempFeature(ID, ModelID, ListOrder, TypeEnglish, DescEnglish) VALUES ({}, '{}', {}, '{}', '{}');".format(self.temp_id, self.maxModelID, self.spec_list_order, self.spec_name_english, self.spec_desc_english)

        print(sql_insert, file=self.dbDetail_file)

    def emitModelTypeMapSQL(self):
        sql_insert = "INSERT INTO ModelTypeMap(ModelID, TypeID) VALUES({}, {});".format(self.maxModelID, 0)
        print(sql_insert, file=self.dbModelMap_file)
        
    def hasModelSQLBeenEmitted(self):
        return self.oneTimePerSheet

    def hasNotBeenEmitted(self):
        self.oneTimePerSheet = True

    def hasBeenEmitted(self):
        self.oneTimePerSheet = False

class ColumnZeroZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content
        mysql_object.incrMaxModelID()

    def get(self):
        return self.myName

class ColumnOneZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnTwoZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)

    def get(self):
        return self.myName

class ColumnThreeZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnFourZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnFiveZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnSixZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnSevenZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnEightZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnNineZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnTenZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName

class ColumnElevenZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

class ColumnTwelveZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

class ColumnThirteenZero():
    myName = ""
    def action(self, cell_content, mysql_object):
        print(cell_content)
        self.myName = cell_content

    def get(self):
        return self.myName
        
 
class ColumnZero():
    def action(self, cell_content, map, mysql_object):
        print("0:", cell_content)
        mysql_object.setSpreadSheetId(cell_content)
        mysql_object.incrTempId()

class ColumnOne():
    def action(self, cell_content, map, mysql_object):
        print("1:", cell_content)
        map[ColumnOneZero().get()] = cell_content

class ColumnTwo():
    def action(self, cell_content, map, mysql_object):
        print("2:", cell_content)
        map[ColumnTwoZero().get()] = cell_content
        if(mysql_object.getSpreadSheetId() == 1):
            mysql_object.setModelName(cell_content)

class ColumnThree():
    def action(self, cell_content, map, mysql_object):
        print("3:",cell_content)
        map[ColumnThreeZero().get()] = cell_content
        mysql_object.setNameEnglish(cell_content)

class ColumnFour():
    def action(self, cell_content, map, mysql_object):
        print("4:",cell_content)
        map[ColumnFourZero().get()] = cell_content

class ColumnFive():
    def action(self, cell_content, map, mysql_object):
        print("5:",cell_content)
        map[ColumnFiveZero().get()] = cell_content

class ColumnSix():
    def action(self, cell_content, map, mysql_object):
        print("6:",cell_content)
        map[ColumnSixZero().get()] = cell_content
        mysql_object.setDescEnglish(cell_content)

class ColumnSeven():
    def action(self, cell_content, map, mysql_object):
        print("7:", cell_content)
        map[ColumnSevenZero().get()] = cell_content
        mysql_object.setSpecOrder(cell_content)

class ColumnEight():
    def action(self, cell_content, map, mysql_object):
        print("8:",cell_content)
        map[ColumnEightZero().get()] = cell_content
        if(mysql_object.getSpreadSheetId() == 1):
            mysql_object.setModelNumber(cell_content)

class ColumnNine():
    def action(self, cell_content, map, mysql_object):
        print("9:",cell_content)
        map[ColumnZeroZero().get()] = cell_content
        if(mysql_object.getSpreadSheetId() == 1):
            mysql_object.setModelShortNumber(cell_content)

class ColumnTen():
    def action(self, cell_content, map, mysql_object):
        print("10:",cell_content)
        map[ColumnZeroZero().get()] = cell_content

class ColumnEleven():
    def action(self, cell_content, map, mysql_object):
        print("11:",cell_content)
        map[ColumnZeroZero().get()] = cell_content

class ColumnTwelve():
    def action(self, cell_content, map, mysql_object):
        print("12:",cell_content)
        map[ColumnZeroZero().get()] = cell_content
        if(mysql_object.getSpreadSheetId() == 1):
            mysql_object.setModelListOrder(cell_content)

class ColumnThirteen():
    def action(self, cell_content, map, mysql_object):
        print("13:",cell_content)
        map[ColumnZeroZero().get()] = cell_content
    
class DoSpreadSheet:
    action_zero = {'0': ColumnZeroZero(), 
                   '1': ColumnOneZero(),
                   '2': ColumnTwoZero(),
                   '3': ColumnThreeZero(),
                   '4': ColumnFourZero(),
                   '5': ColumnFiveZero(),
                   '6': ColumnSixZero(),
                   '7': ColumnSevenZero(),
                   '8': ColumnEightZero(),
                   '9': ColumnNineZero(),
                   '10': ColumnTenZero(),
                   '11': ColumnElevenZero(),
                   '12': ColumnTwelveZero(),
                   '13': ColumnThirteenZero()
                   }
    action_non_zero = {'0': ColumnZero(), 
                   '1': ColumnOne(),
                   '2': ColumnTwo(),
                   '3': ColumnThree(),
                   '4': ColumnFour(),
                   '5': ColumnFive(),
                   '6': ColumnSix(),
                   '7': ColumnSeven(),
                   '8': ColumnEight(),
                   '9': ColumnNine(),
                   '10': ColumnTen(),
                   '11': ColumnEleven(),
                   '12': ColumnTwelve(),
                   '13': ColumnThirteen()
                   }

    rows = {}
    columns = {}

    def __init_(self):
        self.columns = {}
       
        

    def builder(self, row, column, value, mysql_object):
        if(row == 0):
            if(mysql_object.hasModelSQLBeenEmitted()):
                self.columns[str(row)] =  self.rows
                # put this last row into the column map
                mysql_object.emitModelSQL()
                mysql_object.emitModelTypeMapSQL()
                mysql_object.hasBeenEmitted()
            self.rows = {}
            b_object = self.action_zero[str(column)]
            b_object.action(value, mysql_object)
        else:
            b_object = self.action_non_zero[str(column)]
            b_object.action(value, self.rows, mysql_object)
            
            
        
    
    def doOneSpreadSheet(self, spreadSheetPath):
        mySqlClient = MySqlClient();
        mySqlClient.queryMaxModelID()
        print('MaxModelID:', mySqlClient.getMaxModelID())
        workbook = xlrd3.open_workbook(spreadSheetPath)
        sheetNames = workbook.sheet_names()
        for sheetName in sheetNames:
            worksheet = workbook.sheet_by_name(sheetName)
            num_rows = worksheet.nrows - 1
            num_cells = worksheet.ncols - 1
            curr_row = -1
            mySqlClient.hasNotBeenEmitted()
            while curr_row < num_rows:
                curr_row += 1
                row = worksheet.row(curr_row)
                # print('Row:', curr_row)
                curr_cell = -1
                while curr_cell < num_cells:
                    curr_cell += 1
                    # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
                    cell_type = worksheet.cell_type(curr_row, curr_cell)
                    cell_value = worksheet.cell_value(curr_row, curr_cell)
                    print("CurrentSheet", sheetName, "CurrentRow:", curr_row, "CurrentCell:", curr_cell,"CellType:", cell_type, "CellValue:", cell_value)
                    
                    self.builder(curr_row, curr_cell, cell_value,  mySqlClient)
                if(curr_row != 0):
                    mySqlClient.emitDetailSQL()
        mySqlClient.cleanClose()

if __name__ == '__main__':
    doSpreadSheetFeatures = DoSpreadSheet()
    doSpreadSheetFeatures.doOneSpreadSheet("/home/jredden/workspace/database/trunk/sql/moto_product_lines/py/data/2013_Oem_Model_Features.xls")
