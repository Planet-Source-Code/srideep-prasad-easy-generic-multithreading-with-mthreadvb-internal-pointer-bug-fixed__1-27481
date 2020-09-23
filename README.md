<div align="center">

## Easy Generic Multithreading with MThreadVB (Internal Pointer bug fixed)<br/>by Srideep Prasad

</div>

### Description

MThreadVB is a generic multithreader, allowing you to multithread any function or sub. To find out more, read on !

### Source Code

<html>
<head>
<meta http-equiv="content-language" content="en-us">
<meta http-equiv="content-type" content="text/html; charset=windows-1252">
<meta name="generator" content="microsoft frontpage 4.0">
<meta name="progid" content="frontpage.editor.document">
<title>important notice</title>
</head>
<body>
<p><b>mthreadvb - the easy way to multithread !</b></p>
<p>mthreadvb is a generic multithreader for vb - to which i have been making a
few changes here and there.... but it seems that in one of my updates i had
inadvertently referenced an independent dll called vbconsole.dll and had
forgotten to remove it.&nbsp; (this was done for testing and experimentation
purposes)...i had also forgotten to remove an invalid object variable
reference.... as a result, the update may not have worked.... i apologize for
any inconvenience and those of you who had downloaded the buggy code can
download the updated version now !&nbsp; plus this new update has quite a few
more features (and took quite some time to add too !)</p>
<p><b><font color="#000080"><u>fixes / enhancements</u></font></b></p>
<p>1&gt;the vbconsole.dll reference problem has been fixed....</p>
<p>2&gt;now defines a new property objectinthreadcontext, that returns the
reference to the parent object containing the multithreaded sub in context to
the new thread</p>
<p>3&gt;with this, you can now implement file i/o and show forms (though i do
not very much recommend showing forms from multithreaded procedures), from multithreaded subs
(the form show bug was reported by robin lobel - special thanks to him for doing
so !)</p>
<p>4&gt;some users it seems are having problems showing forms within
multithreaded procedures. therefore i have updated the code to actually
demonstrate how to actually show forms from multithreaded procedures....</p>
<p>5&gt;a serious pointer dereferencing bug was causing problems when the
multithreaded sub had a relatively big name. this has now been fixed !</p>
<p>here is the link to the bug fixed code -</p>
<p><a href="http://planet-source-code.com/vb/default.asp?lngcid=26900&amp;lngwid=1">http://planet-source-code.com/vb/default.asp?lngcid=26900&amp;lngwid=1</a>&nbsp;
</p>
<p>do not hesitate to mail be if you notice some bug or problem....</p>
<p>please remember that many of the enhancements were made possible due to
feedback from people at psc.... please continue to give your feedback regarding
any problems in the functioning of mthreadvb</p>
<p><b><u>note: if you consider this code worth your vote, please vote for the main
page by clicking the link above... or register it by clicking <a href="http://www.planet-source-code.com/vb/scripts/voting/voteoncoderating.asp?lngwid=1&amp;txtcodeid=26900&amp;optcoderatingvalue=5">here
!</a></u></b></p>
</body>
</html>

