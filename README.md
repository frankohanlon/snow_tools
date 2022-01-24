# snow_tools
Variety of scripts and such for processing a couple types of snow depth data

## Tool One: Reformat Magnaprobe snow depth data into a Google Earth KML
Pretty simple.  Used for nightly mapping snow course data in order to plan for the next day's course.  Also useful as a quick tool for visualizing the measurement time & ID for corellating with corrections that might be logged in a fieldbook.

## Tool Two: Libre Office sample spreadsheet and macro script for manually processing Campbell Scientific SR50A Snow Depth data.
This tool will probably expand to a python script at some point.  The spreadsheet sample here is data from a site.  Necessary elements:

Column A: Date & Time

Column B: Snow depth data corrected for air temeprature

Column C: Optional but recommended: the quality value

The macro basic file needs to be imported into the LibreOffice Basic editor.  Then run the the Main subroutine.  Automated corrections are added to column D but cleared before every run.  Manual corrections can be made in Column E.  Samples for this column are (all case sensitive):

start of season  : a flag to tell the output program to start the output CSV from this row.

add : tell the program to include this value even if the automated processing would delete it.

delete : tell the program to delete this value even if the automated processing would include it.

override : one of the user adjustable variables is 'MaxFlat' which is used to stop the program if the filtered value is 

end of season : a flag to tell the output program to stop exporting data at this row.

Beyond these items there are some settings that can be adjusted at the top:

Cell B2:  Float value to correct the zero snow distance value to zero.  It's useful to iterate once or twice on this value.  Step one is run through the data with a height of zero.  Then, while viewing the data, identify or estimate the time & date of the start of the snow season. At that time stamp, the snow depth value should be used as the offset.

Cell B3: Set the MaxFlat value (integer).  If you want the program to march through the data start of season to end of season use a larger number.  If you want the program to automatically stop during snow events then try a smaller number.  50 is a good large number and 10 is a good small number.

Cell B4: FLoat value for lower limit acceptable snow depths.  Use this to help truncate out of range data such as when the sensor return scatters off something funny.  In rime ice conditions the distance echo can come from all over and readings can be returned out of expected range.  This value is automatically adjusted down ten centimeters maybe for late winter when frost jacking of the instrument or vegetative surface being compressed over the winter season either of which can affect final snow-free depth.

Cell B5: Float value for upper limit.  Typically this is the height of the snow stand minus the sensor blanking distance (roughly).  

Cell B6: Float value for MaxDelta. This is the maximum allowable change from one reading to the next.  Larger than this the data point is skipped.  However, applies only to marginal confidence quality value measurments.  At the moment all high confidence (low quality values) are included in the final dataset.  Typically this works out ok though data set is adjusted for outliers using the add/delete flags on occasions where the higher confidence quality is in error.

Cell B7: Start row (integer).  Starting row for the data proccessing.  Base 0 not 1 so value of 15 corresponds to row 16.  Typically the first two runs the startrow is 15 and then as the data is processed this will increase due to where in the season the macro terminates.

Cell B8: End Row (integer).  This is an equation that counts the non empty cells in column C.  Not typically changed, just a reference.

Cell B9: Run Stop (integer).  Default is to set this equal to B8.  But, often during analysis it might be StartRow B7 plus 500 or 1000 or something like that if data is being viewed incrementally.

Cell E3: String value for the root name of the file (5min & 60min & csv are appended during output).  Station namme, location, years of winter are typical with words separated by underscore/hyphen.

Cell E4: Output directory for the csv files

Cell E5: Start of season date (just for user's info, this isn't reference by the macro)

Cell E6: End of season date (again, just for user's info)



## Tool 3: Python script for exporting to csv the final data products from the spreadsheet
Reads through the spreadsheet quickly and scoops out the proper data and saves to csv.  There is/was a similar routine in the macro file but it was super slow and a bit wonky on output due to constraints of the macro language.





