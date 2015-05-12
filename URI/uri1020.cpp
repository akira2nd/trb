#include <stdio.h>
#include <conio.h>

main()
{
	int N,d,m,a;
	scanf("%d", &N);
	a = N/365;
	m = (N-(365*a))/30;
	d = N-((a*365)+(m*30));
	printf("%d ano(s)\n", a);
	printf("%d mes(es)\n", m);
	printf("%d dia(s)\n", d);
	
	getch();
	return 0;
}