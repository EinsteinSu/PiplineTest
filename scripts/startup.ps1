#write a file to c:\\installer

New-Item -Path "c:\\installer" -Name "testfile1.txt" -ItemType "file" -Value "This is a text string." -Force