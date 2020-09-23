<div align="center">

## HyperApp \- Write binary code with HTML\!

<img src="PIC2001111181433795.jpg">
</div>

### Description

Create GUIs based entirely on HTML or write machine-specific code for web pages!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-11-01 18:11:44
**By**             |[Jonathan Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-smith.md)
**Level**          |Intermediate
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 6\.0, VB Script
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[HyperApp \-326271112001\.zip](https://github.com/Planet-Source-Code/jonathan-smith-hyperapp-write-binary-code-with-html__1-28576/archive/master.zip)





### Source Code

<p align="center"><i><font size="6"><b>HyperApp HTML Interfacing</b></font></i></p>
<hr>
<p align="left">Special thanks to: Stephan (<a href="http://vbpoint.cjb.net">http://vbpoint.cjb.net</a>)
and Chris Kesler</p>
<p align="left"><b>What Is HyperApp?</b></p>
<blockquote>
 <p align="left">HyperApp is an script-driven, object-oriented library which
 allows you to add binary code to any web page you create. I designed HyperApp
 originally to allow for clean-looking page forms in my applications with
 little work, but soon after discovered many more uses.</p>
 <p align="left">HyperApp has three required dependencies:</p>
 <ul>
  <li>
   <p align="left">Microsoft Script Control (found on Windows 2000 and above,
   or at <a href="http://msdn.microsoft.com/scripting">http://msdn.microsoft.com/scripting</a>)</li>
  <li>
   <p align="left">Microsoft Internet Controls</li>
  <li>
   <p align="left">Microsoft HTML Object Library</li>
 </ul>
</blockquote>
<p align="left"><b>HyperApp for the Layman</b></p>
<blockquote>
 <p align="left">HyperApp is very easy to use. Simply add a reference to 'HyperApp
 HTML Interfacing Object Library 1.0' and add a web browser component. Write
 any code you want to give the page access to in a class file and pass a new
 instance of the class to the HyperApp object. You can even pass forms or any
 other object.</p>
</blockquote>
<p align="left"><b>Accessing HyperApp through a page</b></p>
<blockquote>
 <p align="left">When creating the HTML for your interface, script commands can
 be called by preceding any navigational object with 'happ://' (as opposed to
 'http://'). Immediately following happ://, type the statement you wish to
 call. For example, if you had an object named MyObject, and you wanted to
 access its function OpenFile, you might use the following convention:</p>
 <blockquote>
  <p align="left"><font face="Courier New" size="2">happ://MyObject.OpenFile
  &quot;c:\readme.txt&quot;</font></p>
 </blockquote>
 <p align="left">&nbsp;<i>Note:&nbsp; Any references made like this should be
 encoded with hexidecimal to read something like '</i><i>happ://MyObject.OpenFile%20%22c:/readme.txt%22'.
 Using an HTML editor, such as FrontPage, will automatically encode the links
 to this &quot;web-safe&quot; format. HyperApp will automatically decode these
 hex-encoded URLs.</i></p>
</blockquote>
<p align="left"><b>Reminder: This is Alpha Work</b></p>
<blockquote>
 <p align="left">By releasing this code, I'm not saying it's 100% bug-free.
 However, any bugs you come across, please let me know so I can continue to
 update this, what I hope will be, useful tool.</p>
 <p align="left">Soon to come: a HyperApp plugin for Internet Explorer, which
 will run HyperApp-enabled web pages</p>
 <p align="center">Please vote if you like this code!</p>
 <p align="left">&nbsp;</p>
</blockquote>

