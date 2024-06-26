import uno
import pycomm3


def populateTags(*args):
    #reads tags of a particular type from a PLC and populates the list on a sheet
    #used when connecting to a PLC and you dont have a list of the PAX tags tag type is read from the tab name

    #get config info 
    getConfigInfo()
    from pycomm3 import LogixDriver
    with LogixDriver(CLX_IP) as plc:

        #this  filters a tag based on a type
        clx_tags = [
            tag
            for tag, _def in plc.tags.items()
            if (_def['data_type_name'].lower()) == str(TagType).lower() #case sensitive handling
        ]
        
        #write tags to TagName Column
        for index, plctag in enumerate(clx_tags):
            oSheet.getCellByPosition(tag_Start_Col,tag_Start_Row+index).setString(plctag)
    return

def readtags(*args):
    #this works to read one aoi at a time, migh tbe better to read the whole struct at once but I didnt feel like doing the syntax
   
    getConfigInfo()

    getEndOfSheet()

    #build list of tag parameters 
    parameterlist=[] 
    for param in range(param_name_Start_Col,sheet_last_Col):
        label1 = oSheet.getCellByPosition(param,param_name_Start_Row).getString()
        label2 = oSheet.getCellByPosition(param,param_name_Start_Row+1).getString()
        if label1 == '': break #simple error handling
        if label2 == '': break
        parameterlist.append( label1 + label2 ) #smash 2 cells into one string  
    

    from pycomm3 import LogixDriver
    from time import sleep
    with LogixDriver(CLX_IP, init_tags=True) as plc:
        TagList = [] # List of tags for bulk read - bulk read seemed to cause less hangups when reading a lot of tags

        #find tag list (veritcal TagName column on sheet)
        for index_row, tags in enumerate(range(tag_Start_Row,sheet_last_Row),tag_Start_Row):
            TagName = oSheet.getCellByPosition(tag_Start_Col,tags).String
            if TagName == '' : break # stop reading if it sees no tag name

            #add parameters to Tagname
            for index_col, elem in enumerate(parameterlist): #build a list of tags suffixed with parameters
                Tag = TagName + elem #build tag to read
                TagList.append(Tag) #Create array of the values - probably dont need to do this, just push teh value into the cell

            #read tags+Tagnames
            TagsRead=plc.read(*TagList)

            #populate tag values into cells
            for index_col, val in enumerate(TagsRead):
                _tagval=TagsRead[index_col]
                oSheet.getCellByPosition(param_Start_Col+index_col,index_row).setString(str((_tagval).value))

            TagList.clear()

    return
    
def writetags(*args):
    #this works to write one aoi at a time, might be better to read the whole struct at once but I didnt feel like doing the syntax
    #Suggest reading first, don't want to spend the time writing logic to exlcude emptycells    
    getConfigInfo()

    getEndOfSheet()


    #build list of tag parameters 
    parameterlist=[] 
    for param in range(param_name_Start_Col,sheet_last_Col):
        label1 = oSheet.getCellByPosition(param,param_name_Start_Row).getString()
        label2 = oSheet.getCellByPosition(param,param_name_Start_Row+1).getString()
        if label1 == '': break #simple error handling
        if label2 == '': break
        parameterlist.append( label1 + label2 ) #smash 2 cells into one string  
    

    from pycomm3 import LogixDriver
    from time import sleep
    with LogixDriver(CLX_IP, init_tags=True) as plc:
        TagList = [] # tuple of tags for bulk read - bulk read seemed to cause less hangups when reading a lot of tags
        ValueList =[] #values to be written
        #find tag list (veritcal TagName column on sheet)
        for index_row, tags in enumerate(range(tag_Start_Row,sheet_last_Row),tag_Start_Row):
            TagName = oSheet.getCellByPosition(tag_Start_Col,index_row).String
            if TagName == '' : break # stop reading if it sees no tag name

            #add parameters to Tagname, and add value to value list
            for index_col, elem in enumerate(parameterlist): #build a list of tags suffixed with parameters
                Tag = TagName + elem #build tag to read w/ parameters
                TagList.append(Tag)
                ValueList.append(oSheet.getCellByPosition(param_Start_Col+index_col,index_row).getString())
            
            #oSheet.getCellByPosition(param_Start_Col,index_row).setString(str(ValueList))
            WriteList = list(zip(TagList,ValueList))
            plc.write(*WriteList)

            ValueList.clear()
            WriteList.clear()
            TagList.clear()

    return

def getConfigInfo():
    #helper function for cleanliness
    global doc, setupSheet, CLX_IP,tag_cell,param_Cell, param_Name_Cell
    global oSheet, TagType, _tag_cell, tag_Start_Row, tag_Start_Col, _param_cell
    global param_Start_Row, param_Start_Col, param_name_Start_Row, param_name_Start_Col

    doc = XSCRIPTCONTEXT.getDocument() #gets open document name/reference
    setupSheet = doc.Sheets.getByName("Setup") 
    CLX_IP = setupSheet.getCellRangeByName("CLX_IP").String #get IP address from cell

    tag_cell = setupSheet.getCellRangeByName("Tag_Cell").String #get location of where to start entering/or reading tags
    param_Cell = setupSheet.getCellRangeByName("param_Cell").String #
    param_Name_Cell = setupSheet.getCellRangeByName("param_name_Cell").String #

    oSheet = doc.getCurrentController().getActiveSheet() #open sheet
    TagType = oSheet.Name #use the sheet/tab name for the tag to lookup

    _tag_cell = oSheet.getCellRangeByName(tag_cell)
    tag_Start_Row = _tag_cell.CellAddress.Row
    tag_Start_Col = _tag_cell.CellAddress.Column

    _param_cell = oSheet.getCellRangeByName(param_Cell)
    param_Start_Row = _param_cell.CellAddress.Row
    param_Start_Col = _param_cell.CellAddress.Column  

    _param_name_cell = oSheet.getCellRangeByName(param_Name_Cell)
    param_name_Start_Row = _param_name_cell.CellAddress.Row
    param_name_Start_Col = _param_name_cell.CellAddress.Column    

    return

def getEndOfSheet():
    global sheet_last_Col, sheet_last_Row

   #find end of sheet
    Curs = oSheet.createCursor()
    Curs.gotoEndOfUsedArea(True)
    sheet_last_Col = Curs.Columns.Count 
    sheet_last_Row = Curs.Rows.Count 
    return
