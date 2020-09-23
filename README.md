<div align="center">

## vbFast \- make vb create strings AT LEAST 150 TIMES faster\!

<img src="PIC20021230223405878.JPG">
</div>

### Description

This tutorial shows how to dynamicly create Visual Basic strings up to 150 times faster by calling the OLE Automation library directly. Please vote or leave a comment.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-01-24 11:22:54
**By**             |[jbay101](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jbay101.md)
**Level**          |Advanced
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[vbFast\_\-\_m15208712302002\.zip](https://github.com/Planet-Source-Code/jbay101-vbfast-make-vb-create-strings-at-least-150-times-faster__1-42020/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Visual Basic vs</title>
</head>
<body>
<p><font face="Verdana" size="2"><b>Visual Basic vs. C++</b></font></p>
<p><font face="Verdana" size="2">Visual Basic stores it's strings in a type
referred to in C++ as a BSTR. This type is completely different from the C char
type, as a BSTR doesn't necessarily terminate with a null, and it has a
different header. The C char is stored as an array of bytes, terminating at a
null byte or character 0x0. Unlike C or C++, when you create a string in VB it
is automatically filled with data. </font></p>
<p><font face="Verdana" size="2"><b>The SLOW way  - Visual Basic's String
creation<br>
</b>When you dynamically create a string in Visual Basic, there are only two
methods that VB supports. These are:<br>
1. Using the String function<br>
    <u>Example:<br>
</u></font><font size="2" face="Courier New">   
<font color="#000080">Dim</font> strData <font color="#000080">As String   
</font><font color="#008000">'our string variable</font><br>
    <font color="#000080">Open</font> "test.bin"
<font color="#000080">For Binary Access Read As</font> #1   
<font color="#008000">'open a file</font><br>
    strData = String(LOF(1), 0)   
<font color="#008000">'create a buffer</font><br>
    <font color="#000080">Get</font> #1, , strData   
<font color="#008000">'read data into the buffer</font><br>
    <font color="#000080">Close</font> #1   
<font color="#008000">'close the file</font><br>
</font><font face="Verdana" size="2">    The String function
takes two parameters, the length of the string and the character to fill the
string with.</font></p>
<p><font face="Verdana" size="2">2. Using the Space function<br>
    This is much like using the String function, except it
automatically fills the string with spaces.</font></p>
<p><font face="Verdana" size="2">Now, for the example above all we want is an
empty storage space to fill with data. But VB doesn't do this. In both
instances, VB fills the string with data, which can take a lot of time. This is
where the API optimization comes into play.</font></p>
<p> </p>
<p><font face="Verdana" size="2"><b>The FAST way - the OLE Automation library<br>
</b>The OLE Automation library provides support, not only for the BSTR type but
also for all variable-related operations. To increase the speed of the string
creation, we want to tell the OLE Automation library to create a region of
memory that we can access - without filling it with data. To do this we will use
two functions, RtlMoveMemory in the windows kernel and SysAllocStringByteLen is
the OLE Automation library. The declarations are below.</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Declare Sub</font>
RtlMoveMemory <font color="#000080">Lib</font> "kernel32" (dst<font color="#000080">
As Any</font>, src<font color="#000080"> As Any</font>, <font color="#000080">
ByVal</font> nBytes&)<br>
<font color="#000080">Declare Sub</font> SysAllocStringByteLen&
<font color="#000080">Lib</font> "oleaut32" (<font color="#000080">ByVal</font>
olestr&, <font color="#000080">ByVal</font> BLen&)</font></p>
<p><font face="Verdana" size="2">The RltMoveMemory function copies nBytes bytes
from the src address to the dst address. The SysAllocStringByteLen allocates
BLen of storage space for a BSTR, or in this case a Visual Basic String. In
reality, the Visual Basic String is nothing more than a pointer, or a reference
to an address in memory that can be used to store the data. With this in mind,
we can create out own string allocation function, as shown below.</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Public Function
</font>AllocString(ByVal lSize <font color="#000080">As Long</font>)
<font color="#000080">As String</font><br>
RtlMoveMemory <font color="#000080">ByVal</font> <font color="#000080">VarPtr</font>(AllocString_ADVANCED),
SysAllocStringByteLen(0&, lSize + lSize), 4&<br>
<font color="#000080">End Function</font><br>
<br>
</font><font face="Verdana" size="2">This may look a bit complicated at first
but it is really relatively simple. The function allocates the space and then
copies the 4 byte pointer from this space to the string returned by the
function. If we were to expand the function a little it would look like this:</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Public Function
</font>AllocString(ByVal lSize <font color="#000080">As Long</font>)
<font color="#000080">As String</font><br>
<font color="#000080">Dim</font> lPtr <font color="#000080">As Long   
</font><font color="#008000">'the address of the allocated memory</font><br>
<font color="#000080">Dim</font> lRetPtr<font color="#000080"> As Long   
</font><font color="#008000">'the pointer to the return variable<br>
</font><font color="#000080">Dim</font> sBuffer <font color="#000080">As String   
</font><font color="#008000">'the variable to return</font><br>
lRetPtr<font color="#000080"> = VarPtr</font>(sBuffer)   
<font color="#008000">'the pointer to the string buffer</font><br>
lPtr = SysAllocStringByteLen(0&, lSize + lSize)   
<font color="#008000">'allocate the memory and get it's pointer</font><br>
RtlMoveMemory <font color="#000080">ByVal</font> lRetPtr, lPtr, 4&   
'<font color="#008000">copy the pointer address</font><br>
AllocString = sBuffer    <font color="#008000">'return the string
with the modified pointer</font><br>
<font color="#000080">End Function</font><br>
 </font></p>
<p><font face="Verdana" size="2">It really is not that difficult, and it makes a
HUGE speed increase. This article comes with the above function and a benchmark
to show the dramatic speed difference. Please leave a comment or vote!</font></p>
</body>
</html>

