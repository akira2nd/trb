#include <stdio.h>
#include <conio.h>

main()
{
	int N,I,R;
	scanf("%d", &N);
	printf("%d\n", N);
	I = N/100;
	R = N%100;
	printf("%d nota(s) de R$ 100,00\n", I);
	I = R/50;
	R = R%50;
	printf("%d nota(s) de R$ 50,00\n", I);
	I = R/20;
	R = R%20;
	printf("%d nota(s) de R$ 20,00\n", I);
	I = R/10;
	R = R%10;
	printf("%d nota(s) de R$ 10,00\n", I);
	I = R/5;
	R = R%5;
	printf("%d nota(s) de R$ 5,00\n", I);
	I = R/2;
	R = R%2;
	printf("%d nota(s) de R$ 2,00\n", I);
	I = R/1;
	R = R%1;
	printf("%d nota(s) de R$ 1,00\n", I);
	getch();
	return 0;
}