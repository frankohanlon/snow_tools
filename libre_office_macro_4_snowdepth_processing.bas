option explicit


 REM  *****  BASIC  *****
' Procedure here:
' 1) work through the data staring with 'sub main'.
' 1a look over the data in origin going in 2-5 day increments through the season.
'    add to teh wiki the start, end of snow season and the offset for this data set.
' 2) once season is complete hope down to 'sub datacopy' and run.  This will pull out all the hourly data and put it into a new column.
' ---3)--- final 2 steps: run 'sub output_to_csv_5min' and 'sub output_to_csv_60min'
' 3) To output data in final form, run:
      '  /var/site/uaf_sp/bin/teller_top_snow_depth_output.py   
' 1):
Sub Main

    dim columnDate As Integer
    dim columnrawDist As Integer     
    dim columnQuality As Integer
    dim columnAUTO As Integer	
    dim columnMANUAL As Integer
    dim columnSDC As Integer
    dim columnFILT As Integer
    dim columnINFIL As Integer
    dim column1hrMA As Integer
    dim column3hrMA As Integer
'    dim column3hrQual As Integer
    dim columnQualFilterOutput As Integer
    dim columnQualNumOutput as integer
    dim columnQualFilterMarginal as integer
    dim startrow As Long
	dim lastrow as long
    Dim currow As Long 
    Dim timestamp                   ' current time & date
    Dim curdistval As Double        ' distance value from spreadsheet
    Dim curqualval As Double        ' current quality value
    Dim diffqualval As Double       ' difference from 3hr MA quality value
    Dim snowdepth As Double         ' pt1 snow depth value
    Dim prevSDCval                  ' previous time step snowdepth converison value
    Dim curSDCval                   ' snow depth conversion value from spreadsheet
    Dim futureSDCval As Double      ' future snow depth conversion value from spreadsheet
    Dim previnfilval as double      ' previous time step infilled value
    Dim curinfilval As Double       ' current infilled value
    Dim nextinfilval As Double      ' future infilled value
    Dim cur3hrMAval As Double       ' current 3 hour moving average value
    Dim manualadd As Boolean
    Dim manualdelete As Boolean
    Dim temp1, temp2, badloop
    Dim tempstring
    Dim RunStop as long
'    dim LowQOverride as boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Set the spreadsheet related  variables ''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    columnrawDist = 1 ' column B = Raw Depth
    columndate  = 0    ' column A = Date
    columnQuality = 2 ' column C = Quality #
    columnAUTO = 3    
    columnMANUAL = 4     ' column E = Manual Add
    columnSDC  = 5     ' column F = Raw Snow Depth
    columnFILT = 6    ' column G = Filtered Snow Depth
    columnINFIL = 7   ' column H = Infilled & Filtered Snow Depth
    column1hrMA = 8   ' column I = 1 hour moving average
    column3hrMA = 9  ' column J = 3 hour moving average
'    column3hrQual = 10 ' column K = 3 hour moving average of quality number
    columnQualFilterOutput = 10
    columnQualFilterMarginal = 11
    columnQualNumOutput = 12
	startrow = get_SS_value(1,6)        ' starting row for all computations
	lastrow = get_SS_value(1,7)         ' last row for all computations
	RunStop = get_SS_value(1,8)         ' intermediate row for all computations
'	lowqoverride = get_SS_value(1,9)  
	
'	my_doc = ThisComponent
'	my_sheets = my_doc.Sheets 
    set_SS_string(6,0,"--RUNNING--")
    '''''''''''''''''''''''''''''''''''
    ''''' snow algorithm variables  '''
    '''''''''''''''''''''''''''''''''''
    dim lowlimit As Double 
    dim springlowlimit as double
    Dim sensorheight As Double
    Dim highlimit As Double
    dim maxDelta As Double
    Dim maxflat As Integer
    Dim infilindex As Long
    Dim autoerror As Long
    Dim manualerror As Long
    dim sd_timedelta as double
	maxflat = get_SS_value(1,2)  ' start with 10
    lowlimit = get_SS_value(1,3)
    sensorheight =  get_SS_value(1,1)
    highlimit = sensorheight + get_SS_value(1,4)    
    infilindex = 0
    autoerror = 0
    manualerror = 0
    springlowlimit = 0
    maxdelta = get_SS_value(1,5)
    
    '''''''''''''''''''''''''''''''''''''
    ''' clear the auto edits column:  '''
    '''''''''''''''''''''''''''''''''''''
    clear_auto_edits(startrow,lastrow)

    '''''''''''''''''''''''''''''''''''''''''''''''''
    '' Now ready to copy over the snow algorithm.  ''
    '''''''''''''''''''''''''''''''''''''''''''''''''

	for currow = startrow to lastrow
		if get_SS_string(columnMANUAL,currow) = "end of season" then  
            set_SS_value(5,0, currow)  ' F1
            set_SS_string(6,0,"--END--")
            set_SS_string(13,currow-1,str(currow))	 
            set_SS_string(8,0,get_SS_string(0,currow))   
            fastforward(currow)        
			exit sub
		endif
		timestamp =  get_SS_String(columndate, currow)
		if int(left(timestamp,2)) < 7 then
		    springlowlimit = -10
		endif   
		''''''''''''''''''''''''''''''''''''''''''''
		'''' Part 1: convert distance to depth. ''''
		'''' and get other inputs from sheet    ''''
		''''''''''''''''''''''''''''''''''''''''''''		
 		curdistval = get_SS_value(columnrawDist,currow)
 		curqualval = get_SS_value(columnQuality,currow)
 '----------------------------------------------------------------------------------------------------
 		snowdepth = curdistval - sensorheight
 '----------------------------------------------------------------------------------------------------
		set_SS_value(columnSDC,currow,snowdepth)  ' place initial value in snow depth column
        curSDCval = snowdepth
        previnfilval = snowdepth
        
        '''''''''''''''''''''''''''''''''''
        '''' Remaining Initialization  ''''
        '''''''''''''''''''''''''''''''''''
        If currow = startrow Then
			set_SS_value(columnSDC,currow,snowdepth)
			set_SS_value(columnFILT, currow, snowdepth)
			set_SS_value(columnINFIL, currow, snowdepth)
        End If

        
		''		'''''''''''''''''''''''''''''
        '''' part 2: filter step 1. '''
        '''''''''''''''''''''''''''''''
        If currow > startrow Then
            
            '''''''''''''''''''''''''''''''''''''''''
            '''' Set Manual add and  delete flags ''''
            ''''''''''''''	'''''''''''''''''''''''''''
			if get_SS_string(columnMANUAL,currow) = "delete" then 
				manualdelete = True 
            	infilindex = 0			
			Else 
				manualdelete = False
			endif
			If get_SS_string(columnMANUAL,currow) = "add" Then 
				manualadd = True 
            	infilindex = 0				
			Else 
				manualadd = False
			endif
			' this isn't quite the right place for this infil but...
            previnfilval = get_SS_value(columnINFIL,currow-1)
            sd_timedelta = Abs(previnfilval - snowdepth)

            
'            if curqualval > 0 and curqualval < 210 then 'Good
'                set_SS_string( columnQualNumOutput, currow, 3.0)
'            elseif curqualval >= 210 and curqualval < 300 then 'marginal
'                set_SS_string( columnQualNumOutput, currow, 2.0)            
'            else 'bad
'                set_SS_string( columnQualNumOutput, currow, 1.0)            
'            endif
            
            If curqualval = 0 Then   ' quality # = 0 means bad SDI12 comms
            	set_SS_string( columnAUTO, currow, "bad quality")
            	set_SS_string( columnSDC, currow, "")
            	set_SS_string( columnFILT, currow, "")
            	set_SS_value( columnINFIL ,currow, previnfilval)
            	set_SS_value( columnQualFilterOutput, currow, get_SS_value(columnQualFilterOutput, currow-1) )
            	infilindex = 0
            	            	
            ElseIf snowdepth <= (lowlimit + springlowlimit) Then  ' errant measurement, delete point.
            
            	set_SS_string( columnAUTO, currow, "low limit")
				' Not removing from SDC column because I'd like to see if the ground surface is set wrong.
            	set_SS_string( columnFILT, currow, "")
				set_SS_value( columnQualFilterOutput, currow, get_SS_value(columnQualFilterOutput, currow-1) )
            	set_SS_value( columnINFIL ,currow, previnfilval)
            	
            ElseIf snowdepth >= highlimit Then  ' errant measurement, delete point.
            	set_SS_string( columnAUTO, currow, "high limit")
				' Not removing from SDC column because I'd like to see if the high limit is set wrong.
            	set_SS_string( columnFILT, currow, "")
				set_SS_value( columnQualFilterOutput, currow, get_SS_value(columnQualFilterOutput, currow-1) )
            	set_SS_value( columnINFIL ,currow, previnfilval)
            	
            ElseIf curqualval < 210 and curqualval > 0 and manualdelete = false then
            	' good value overrite, output to sheet & reset infilindex.
            	set_SS_value( columnFILT, currow, snowdepth)
            	set_SS_value( columnINFIL ,currow, snowdepth)
                infilindex = 0
                set_SS_string(columnAuto, currow, "HQ Q")       	
                set_SS_value( columnQualFilterOutput, currow, snowdepth )
            ElseIf curqualval > 300 and manualadd = false then
            	' bad rated measurement.
            	set_SS_string( columnFILT, currow, "")
            	set_SS_value( columnINFIL ,currow, previnfilval)
                infilindex = 0
                set_SS_string(columnAuto, currow, "LQ Q")       	
                set_SS_value( columnQualFilterOutput, currow, previnfilval )                
            ElseIf SD_timedelta > maxDelta Or manualdelete = True Then
            	set_SS_String(columnAUTO, currow, "maxdelta or manual delete")
                set_SS_string( columnFILT, currow, "")
                if manualdelete = false then
                    set_SS_value( columnAUTO, currow, Abs(previnfilval - snowdepth) )
                endif
            Else
            	' good value, output to sheet & reset infilindex.
            	set_SS_value( columnFILT, currow, snowdepth)
            	set_SS_value( columnINFIL ,currow, snowdepth)
                infilindex = 0
            End If
            ''''''''''''''''''''''''''''' Good Quality Point Series ''''''''''''''''''''''
            If curqualval < 210 and curqualval > 0 then
                set_SS_value( columnQualFilterOutput, currow, snowdepth )
            else:
                set_SS_value( columnQualFilterOutput, currow, get_SS_value(columnQualFilterOutput, currow-1) )            
            endif
            ''''''''''''''''''''''''''Marginal Quality Point Series'''''''''''''
            If curqualval < 300 and curqualval > 0 and snowdepth > lowlimit + springlowlimit and snowdepth < highlimit then
                set_SS_value( columnQualFilterMarginal, currow, snowdepth )
            else:
                set_SS_value( columnQualFilterMarginal, currow, get_SS_value(columnQualFilterMarginal, currow-1) )            
            endif
            
            If manualadd = True Then   ' manually add a datapoint
            	set_SS_value(columnFILT, currow, snowdepth)
            	set_SS_value(columnINFIL, currow, snowdepth)
            	set_SS_value(columnQualFilterOutput, currow, snowdepth)
                curinfilval = snowdepth
                infilindex = 0
            Else
            	' at this point filter column should be filled if the data is good.
                If IsBlank(columnFILT, currow) Then
					set_SS_value(columnINFIL, currow, get_SS_string(columnINFIL, currow - 1) )
					tempstring = get_SS_string(columnQualFilterOutput, currow)
                Else:  ' (IsBlank(columnFILT, currow) = false 0
                	set_SS_value(columnINFIL, currow, snowdepth)
                    curinfilval = snowdepth
                    infilindex = 0
                End If

                If snowdepth <= lowlimit + springlowlimit Or curqualval = 0  Then  ' delete point.
                    set_SS_string(columnAuto, currow, "autodelete")
                    set_SS_string(columnSDC, currow, "")
                    set_SS_value(columnQualFilterOutput, currow, get_SS_value(columnQualFilterOutput, currow - 1 ) )                    
				elseif  snowdepth > highlimit then
                    set_SS_string(columnAuto, currow, "autodelete")
                    set_SS_string(columnSDC, currow, "")
                    set_SS_value(columnQualFilterOutput, currow, get_SS_value(columnQualFilterOutput, currow - 1 ) )                    
					infilindex = 0
                End If
            End If
			' Override in this case refers to the maxdelta / maxflat parameter.	
	        If get_SS_string(columnMANUAL, currow) = "override" or get_SS_string(columnMANUAL, currow) = "flat" Then
	            infilindex = 0
	        End If
	        If infilindex > maxflat or currow > runstop Then
	            'calculate
	            if runstop > 0 then
		            set_SS_value(5,0, currow)  ' F1
	                set_SS_string(6,0,"--END--")
	                set_SS_string(13,currow-1,str(currow))	 
	                set_SS_string(8,0,get_SS_string(0,currow)) 
	                fastforward(currow)          
		            Exit Sub
		        else
		            set_SS_value(5,0, currow)  ' F1
	                set_SS_string(6,0,"--END--")
	                set_SS_string(13,currow-1,str(currow))	 
	                set_SS_string(8,0,"Need to add RunStop value to Cell B9") 
	                fastforward(currow)          
		            Exit Sub
				endif		            
	        End If
			if isBlank(columnFILT,currow) then
				infilindex = infilindex + 1
				set_SS_value(columnINFIL, currow, previnfilval )
			else
				set_SS_value(columnINFIL, currow, snowdepth)
			endif
	        ' part 4: update status
	        If currow Mod 50 = 0 Then
	            set_SS_value(5,0, currow)  ' F1
	            set_SS_string(8,0,get_SS_string(0,currow))
	        End If
		End if
 	next currow
 	set_SS_string(6,0,"--END--")
 	
End Sub  ' END MAIN
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Specialty Functions & Subs   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FUNCTION get_SS_value (column as integer, row as long)
    dim my_doc, my_sheets, my_cell
    my_doc = ThisComponent
    my_sheets = my_doc.Sheets 
    my_cell = ThisComponent.Sheets(0).getCellByPosition(column,row)  ' (column, row)
	get_SS_value =  my_cell.Value	
end FUNCTION

FUNCTION get_SS_string (column as integer, row as long)
    dim my_doc, my_sheets, my_cell
    my_doc = ThisComponent
    my_sheets = my_doc.Sheets 
    my_cell = ThisComponent.Sheets(0).getCellByPosition(column,row)  ' (column, row)
	get_SS_string =  my_cell.String
end FUNCTION

Sub set_SS_value(column as integer, row as long, SS_value)
    dim my_doc, my_sheets, my_cell
    my_doc = ThisComponent
    my_sheets = my_doc.Sheets 
    my_cell = ThisComponent.Sheets(0).getCellByPosition(column,row)  ' (column, row)
	my_cell.Value = SS_value	
end sub
Sub set_SS_string(column as integer, row as long, SS_string)
    dim my_doc, my_sheets, my_cell
    my_doc = ThisComponent
    my_sheets = my_doc.Sheets 
    my_cell = ThisComponent.Sheets(0).getCellByPosition(column,row)  ' (column, row)
	my_cell.String = SS_string
end sub

Function IsBlank(column as integer, row as long)
    dim my_doc, my_sheets, my_cell
    my_doc = ThisComponent
    my_sheets = my_doc.Sheets 
    my_cell = ThisComponent.Sheets(0).getCellByPosition(column,row)  ' (column, row)
    If  (my_cell.Type = com.sun.star.table.CellContentType.EMPTY) then
    	IsBlank = True
    Else
    	IsBlank = False
    end if
end function

sub clear_auto_edits(startrow,endrow)
	rem ---------------------------------- 
	rem select and delte
	dim document, dispatcher
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	dim area_str as string
	area_str = "$D$15:$D$80000"
    ' area_str =  "$D$" & cstr(startrow) & ":$D:$" & cstr(endrow)
    dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = area_str
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
	dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())
end sub

sub fastforward(currow)
	dim document, dispatcher
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	dim area_str as string
	area_str = "$L$" & currow
    dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = area_str
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())

end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Extract finished hourly data from 5 minute product '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub datacopy
    dim columnDate As Integer
    dim columnAUTO As Integer	    
    dim columnMANUAL As Integer
    dim columnSDC As Integer
    dim columnINFIL As Integer
    dim column1hrMA As Integer
    dim column3hrMA As Integer
    dim startrow As Long
	dim lastrow as long
    Dim currow As Long     
    Dim timestamp                   ' current time & date
    Dim timez as date
	dim curminute
	dim outrow
	dim sdcval as double
	dim highlimit as double
	dim snow_start as date
	dim snow_end as date
    Dim manualdelete As Boolean
    snow_start = get_SS_value(4,4)
    snow_end = get_SS_value(4,5)    
	
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Set the spreadsheet related  variables ''''''''
    ''' sub datacopy                           ''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    columndate  = 0    ' column A = Date
    columnAUTO = 3      
    columnMANUAL = 4     ' column E = Manual Add
    columnSDC  = 5     ' column F = Raw Snow Depth
    columnINFIL = 7   ' column H = Infilled & Filtered Snow Depth
    column1hrMA = 8   ' column I = 1 hour moving average
    column3hrMA = 9  ' column J = 3 hour moving average
    
	startrow = get_SS_value(1,6)        ' starting row for all computations
	lastrow = get_SS_value(1,7)         ' last row for all computations
    outrow = 15
    highlimit = get_SS_value(1,4)    
    set_SS_string(14,0, "---Running Hourly Extract Copy---")
	for currow = startrow to lastrow
		timez = get_SS_value(0,currow)
		
		curminute = minute(timez)
		if timez >= snow_start and timez <= snow_end then		
			if curminute = 0 then
				manualdelete = false
   				if get_SS_string(columnMANUAL,currow) = "NAN" or get_SS_string(columnAUTO,currow) = "autodelete" then 
					manualdelete = True 
				endif
			
				' top of the hour... then output.
				set_SS_value(14,outrow,get_ss_value(columnDate,currow))
				sdcval =  get_ss_value(columnSDC, currow)
				if manualdelete = True then
					set_SS_value(15,outrow, 6999)
				else:
					set_SS_value(15,outrow, sdcval)
				endif
				set_SS_value(16,outrow, get_ss_value(columnINFIL, currow) )			
				set_SS_value(17,outrow, get_ss_value(column1hrMA, currow) )
				set_SS_value(18,outrow, get_ss_value(column3hrMA, currow) )						
				outrow = outrow + 1
			endif
		endif			
	next
    set_SS_string(14,0, "---Hourly Extract Stopped---")	
	   
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' EXPORT to CSV  ''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub output_to_csv_5min
    dim columnDate As Integer
    dim columnAUTO As Integer	
    dim columnMANUAL As Integer    
    dim columnSDC As Integer
    dim columnINFIL As Integer
    dim column1hrMA As Integer
    dim column3hrMA As Integer
    dim startrow As Long
	dim lastrow as long
    Dim currow As Long     
    Dim timestamp                   ' current time & date
    Dim timez as date
	dim curminute
	dim sdcval as double
	dim sdcstr as string
	dim infilval as string
	dim MA1hrval as string
	dim MA3hrval as string
	dim highlimit as double
	dim outfile_5min as string
	dim outdirectory as string
	dim outfilehandle_5min
	dim outstring as string
	dim snow_start as date
	dim snow_end as date
    Dim manualdelete As Boolean	
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Set the spreadsheet related  variables ''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    columndate  = 0    ' column A = Date
    columnAUTO = 3  
    columnMANUAL = 4     ' column E = Manual Add    
    columnSDC  = 5     ' column F = Raw Snow Depth
    columnINFIL = 7   ' column H = Infilled & Filtered Snow Depth
    column1hrMA = 8   ' column I = 1 hour moving average
    column3hrMA = 9  ' column J = 3 hour moving average
	startrow = get_SS_value(1,6)        ' starting row for all computations
	lastrow = get_SS_value(1,7)         ' last row for all computations
    highlimit = get_SS_value(1,4)   
    snow_start = get_SS_value(4,4)
    snow_end = get_SS_value(4,5)    
        
    outfile_5min = get_SS_string(4,2) & "_5min.csv"
    outdirectory = get_SS_string(4,3)
    outfile_5min = outdirectory & outfile_5min
    outfilehandle_5min  = FreeFile()  
    open outfile_5min for output as #outfilehandle_5min 
    write #outfilehandle_5min, get_SS_string(1,0)
	write #outfilehandle_5min, "TZ=UTC+0, Snow Depth,Snow Depth: 3hr moving average, Snow Depth: 1hr moving average, Snow Depth: No Filtering"
	write #outfilehandle_5min, ",centimeters,centimeters,centimeters,centimeters"
	write #outfilehandle_5min, "(TZ=UTC+0),Smp,Avg,Avg,Smp"
  
    set_SS_string(14,0, "---Running 5min csv file save---")
	for currow = startrow to lastrow
	    manualdelete = False
        If currow Mod 50 = 0 Then
            set_SS_value(5,0, currow)  ' F1
            set_SS_string(8,0,get_SS_string(0,currow))
        End If	
 		if get_SS_string(columnMANUAL,currow) = "NAN" or get_SS_string(columnAUTO,currow) = "autodelete" then 
			manualdelete = True 
		endif
        
		timez = get_SS_value(0,currow)
		curminute = minute(timez)
		sdcval =  get_ss_value(columnSDC, currow)
		infilval = get_ss_string(columnINFIL, currow) 	
		MA1hrval = get_ss_string(column1hrMA, currow) 
		MA3hrval = get_ss_string(column3hrMA, currow) 						
		if manualdelete = True then
			outstring = format(timez,"yyyy-mm-dd hh:mm:ss") & ",6999,6999,6999,6999"
		else:
			sdcstr =  get_ss_string(columnSDC, currow)
			outstring = format(timez,"yyyy-mm-dd hh:mm:ss") & ","  & infilval & "," & MA3hrval & "," & MA1hrval & "," & sdcstr
		endif		
		if timez >= snow_start and timez <= snow_end then
			write #outfilehandle_5min, outstring
		endif			
		
	next
	close #outfilehandle_5min
	shell "xed " & outfile_5min
	set_SS_string(14,0, "---Hourly CSv Export Stopped---")	

end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub output_to_csv_60min                '''''''''''''''''''''     60 minute  '''
    dim columnDate As Integer
    dim columnAUTO As Integer	
    dim columnMANUAL As Integer
    dim columnSDC As Integer
    dim columnINFIL As Integer
    dim column1hrMA As Integer
    dim column3hrMA As Integer
    dim startrow As Long
	dim lastrow as long
    Dim currow As Long     
    Dim timestamp                   ' current time & date
    Dim timez as date
	dim curminute
	dim sdcval as double
	dim sdcstr as string
	dim infilval as string
	dim MA1hrval as string
	dim MA3hrval as string
	dim highlimit as double
	dim outfile_60min as string
	dim outdirectory as string
	dim outfilehandle_60min
	dim outstring as string
	dim snow_start as date
	dim snow_end as date
    Dim manualdelete As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Set the spreadsheet related  variables ''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    columndate  = 0    ' column A = Date
    columnAUTO = 3  
    columnMANUAL = 4     ' column E = Manual Add    
    columnSDC  = 5     ' column F = Raw Snow Depth
    columnINFIL = 7   ' column H = Infilled & Filtered Snow Depth
    column1hrMA = 8   ' column I = 1 hour moving average
    column3hrMA = 9  ' column J = 3 hour moving average
	startrow = get_SS_value(1,6)        ' starting row for all computations
	lastrow = get_SS_value(1,7)         ' last row for all computations
    highlimit = get_SS_value(1,4)   
    snow_start = get_SS_value(4,4)
    snow_end = get_SS_value(4,5)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Setup File for output ''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	set_SS_string(14,0, "---Running 60min csv file save---")
    outfile_60min = get_SS_string(4,2) & "_60min.csv"
    outdirectory = get_SS_string(4,3)
    outfile_60min = outdirectory & outfile_60min        
    outfilehandle_60min  = FreeFile()  
    open outfile_60min for output as #outfilehandle_60min   
    
	write #outfilehandle_60min, get_SS_string(1,0)
	write #outfilehandle_60min, "TZ=UTC+0, Snow Depth,Snow Depth: 3hr moving average, Snow Depth: 1hr moving average, Snow Depth: No Filtering"
	write #outfilehandle_60min, ",centimeters,centimeters,centimeters,centimeters"
	write #outfilehandle_60min, "(TZ=UTC+0),Smp,Avg,Avg,Smp"
	for currow = startrow to lastrow
		timez = get_SS_value(0,currow)
		curminute = minute(timez)
		if curminute = 0 then
			sdcval =  get_ss_value(columnSDC, currow)
			infilval = get_ss_string(columnINFIL, currow) 	
			MA1hrval = get_ss_string(column1hrMA, currow) 
			MA3hrval = get_ss_string(column3hrMA, currow) 
			manualdelete = false
 		if get_SS_string(columnMANUAL,currow) = "NAN" or get_SS_string(columnAUTO,currow) = "autodelete" then 
			manualdelete = True 
		endif
			if manualdelete = true then
				outstring = format(timez,"yyyy-mm-dd hh:mm:ss") & ",6999,6999,6999,6999"
			else:
				sdcstr =  get_ss_string(columnSDC, currow)
				outstring = format(timez,"yyyy-mm-dd hh:mm:ss") & ","  & infilval & "," & MA3hrval & "," & MA1hrval & "," & sdcstr
			endif		
			' top of the hour... then output.
			if timez >= snow_start and timez <= snow_end then
				write #outfilehandle_60min, outstring
			endif		
		endif
	next	
	
	set_SS_string(14,0, "---Hourly CSv Export Stopped---")	
	close #outfilehandle_60min	    
	shell "xed " & outfile_60min
end sub

' Steps 2 & 3)
sub final_outputs
   ' superseded by new python utility (2021-12)
   ' /var/site/uaf_sp/bin/teller_top_snow_depth_output.py   
   ' important note: add the start & stop dates to E5 and E6 of the spreadsheet to set the export correcctly...
   call output_to_csv_5min
   call output_to_csv_60min
end sub

