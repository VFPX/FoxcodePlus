public VGO_Dat as object
VGO_Dat = newobject("custom")

for x = 1 to 1700
	VGO_Dat.AddProperty(sys(2015),"teste"+sys(2015))
endfor	




* create a long file name in the FoxPro temp drive and directory
   lcTempFileName = SYS(2023) + "\" + "this is a very long filename.txt"
   liHandle = FCREATE(lcTempFileName)
   = FCLOSE(liHandle)

   ? "Original long filename is: " + lcTempFileName
   ? "Short filename is: " + lfn2sfn(lcTempFileName)

   * delete when finished
   DELETE FILE (lcTempFileName)

   FUNCTION lfn2sfn
   *
   * Converts a Win32 long file name to its short file name equivalent
   *
   * passed: long file name, must already exist for this to work
   *
   * returns: fully qualified short file name, or empty string
   * if error is encountered
   *

   PARAMETERS lcInputString

   DECLARE INTEGER GetShortPathName IN Kernel32 STRING @lpszLongPath, ;
      STRING @lpszShortPath, INTEGER cchBuffer
   DECLARE INTEGER GetLastError IN Win32api

   #DEFINE MAX_PATH 255

   * buffer to receive converted file name
   lcOutputString = SPACE(MAX_PATH)

   * length of receiving buffer
   llcbOutputString = LEN(lcOutputString)

   * if successful, llretval will contain the length of the
   * output string
   llretval = GetShortPathName(@lcInputString, @lcOutputString,;
      llcbOutputString)
   IF llretval = 0
   * uncomment for error code
   * wait window "Error occurred, code is: " + ltrim(str(GetLastError()))
      RETURN ""
   ENDIF

   * truncate it at the length the return value indicates
   lcOutputString = LEFT(lcOutputString, llretval)

   RETURN lcOutputString