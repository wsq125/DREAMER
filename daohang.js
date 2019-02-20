// JavaScript Document
function opens(obj)
{
for(var i = 1;i<=6;i++)
{
if(i == obj)
{document.getElementById("video"+i).style.display="block"
}
else
document.getElementById("video"+i).style.display="none"
}
}