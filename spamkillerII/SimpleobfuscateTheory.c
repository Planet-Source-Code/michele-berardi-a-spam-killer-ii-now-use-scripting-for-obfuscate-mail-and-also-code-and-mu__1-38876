/*
C version of the base plain-html obfuscator!

           (C) 2002 Michele Berardi
           http://web.tiscali.it/mberardi
*/

#include <stdio.h>
#include <string.h>
int i;
char *emailaddr = "nospam@nospam.it";
void main()

{
for (i=0;i<strlen(emailaddr);i++){printf ("&#%d;",(short)emailaddr[i]);}
}

