<span style="font-family: Courier New, Courier, monospace;">This script can be used to check and report the status of the Smtp Service and the Virtual SMTP Server created in IIS. For use with Windows 2008+</span>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;">
</span> <span style="font-family: &quot;verdana&quot; , sans-serif;">
</span> <u><span style="font-size: medium;"><b><span style="font-family: &quot;verdana&quot; , sans-serif;">Sample Report Output</span></b></span></u></span>
<div class="separator" style="clear: both; text-align: center;">
</div>
<div class="separator" style="clear: both; text-align: center;">
</div>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;">
</span> </span>
<div class="separator" style="clear: both; text-align: center;">
</div>
<div class="separator" style="clear: both; text-align: center;">
</div>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;"></span>
</span> <div class="separator" style="clear: both; text-align: center;">
</div>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;"><a href="http://2.bp.blogspot.com/-q3QUhe4OrG8/VvvnB61ZxnI/AAAAAAAAAl0/ymeg-mZorGAxeEbGi08vr2A-19A_mC85w/s1600/report_sample.PNG" imageanchor="1"><img border="0" src="https://2.bp.blogspot.com/-q3QUhe4OrG8/VvvnB61ZxnI/AAAAAAAAAl0/ymeg-mZorGAxeEbGi08vr2A-19A_mC85w/s1600/report_sample.PNG" /></a></span> </span>
<div class="separator" style="clear: both; text-align: center;">
</div>
<span style="font-family: Courier New, Courier, monospace;">----------------------------</span>
<div class="separator" style="clear: both; text-align: center;">
</div>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;">
</span> <span style="font-family: &quot;verdana&quot; , sans-serif; font-size: medium;"><u><b><span style="font-family: &quot;trebuchet ms&quot; , sans-serif;">Download</span></b></u>&nbsp;</span></span>
<span style="font-family: Courier New, Courier, monospace;">You can download the script from here<span style="font-size: small;">:</span></span>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;">
</span> </span>
<ul>
<li><span style="font-family: Courier New, Courier, monospace;"><a href="https://github.com/junecastillote/Get-IISSmtpState" target="_blank">Version 1.1 (GitHub)</a></span></li>
<ul>
<li><span style="font-family: Courier New, Courier, monospace;">Removed Local Queue Counter</span></li>
<li><span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;">Removed </span><span style="font-family: &quot;verdana&quot; , sans-serif;">Remote Queue Counter</span></span></li>
<li><span style="font-family: Courier New, Courier, monospace;">Added Queue, Pickup, Drop and BadMail counter</span></li>
<li><span style="font-family: Courier New, Courier, monospace;">Fixed some formatting issues</span></li>
<li><span style="font-family: Courier New, Courier, monospace;">Replaced CSS Color theme (if you prefer the old theme, just copy the $css_string variable from the older version.</span></li>
<li><span style="font-family: Courier New, Courier, monospace;">Some code optimization</span></li>
</ul>
<li><span style="font-family: Courier New, Courier, monospace;">Version 1.0 (GitHub)</span></li>
<ul>
<li><span style="font-family: Courier New, Courier, monospace;">Initial version</span></li>
</ul>
</ul>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;">
</span> <span style="font-family: &quot;verdana&quot; , sans-serif;">To run, no parameters required, just execute the script from PowerShell.</span></span>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;">
</span> <span style="font-size: medium;"><u><b><span style="font-family: &quot;verdana&quot; , sans-serif;">The Variables</span></b></u></span></span>
<span style="font-family: Courier New, Courier, monospace;">Make sure to edit the following variables to conform to your environment or requirements</span>
<span style="font-family: Courier New, Courier, monospace;"><span style="font-family: &quot;verdana&quot; , sans-serif;">
</span> </span>
<div class="separator" style="clear: both; text-align: center;">
</div>
<div class="separator" style="clear: both; text-align: center;">
<span style="font-family: Courier New, Courier, monospace;"><a href="http://3.bp.blogspot.com/-lcOv9kfpywI/VRVytnwYHnI/AAAAAAAAAfE/xbZ2QyY1AJc/s1600/options.gif" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><span style="font-family: &quot;verdana&quot; , sans-serif;"></span></a><span style="font-family: &quot;verdana&quot; , sans-serif;"><a href="http://1.bp.blogspot.com/-MuYncZivwGo/Vvvnjr3hb3I/AAAAAAAAAl8/Bd5-ynIa4vQjkEAfMNpp640NKr5tsi82w/s1600/variables_sample.PNG" imageanchor="1"><img border="0" src="https://1.bp.blogspot.com/-MuYncZivwGo/Vvvnjr3hb3I/AAAAAAAAAl8/Bd5-ynIa4vQjkEAfMNpp640NKr5tsi82w/s1600/variables_sample.PNG" /></a></span></span></div>
<div class="separator" style="clear: both; text-align: center;">
<a href="http://4.bp.blogspot.com/-8Ug6pISa_Rw/VRQTv8lLFXI/AAAAAAAAAew/J9wReBEJW9Q/s1600/variables.gif" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><span style="font-family: Courier New, Courier, monospace;">
</span></a></div>
<span style="font-family: Courier New, Courier, monospace;">
</span> <span style="font-family: &quot;tahoma&quot;; font-size: x-small;"><span style="font-family: &quot;tahoma&quot;; font-size: x-small;"><span style="font-family: &quot;trebuchet ms&quot; , sans-serif;"></span><span style="font-family: &quot;trebuchet ms&quot; , sans-serif;"></span></span>
</span>
