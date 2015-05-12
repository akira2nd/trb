#include <stdio.h>
#include <conio.h>
#include <string.h>
#include <stdlib.h>

main(){
	char *p, *c;
	char a[10];

	scanf("%c", &a);

	*p = *a;

	printf("%c\n", p);
	printf("%s\n", *p);
	//printf("%c\n", *p);
	//printf("%s\n", *p);

	getch();
}