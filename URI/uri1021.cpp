#include <stdio.h>
#include <conio.h>

main()
{
	double N,R,I;
	scanf("%lf", &N);
	printf("NOTAS:\n");
	I = (int)N/100;
	N = N-(I*100);
	R = N;
	printf("%d nota(s) de R$ 100.00\n", (int)I);
	I = (int)N/50;
	N = N-(I*50);
	R = N;
	printf("%d nota(s) de R$ 50.00\n", (int)I);
	I = (int)N/20;
	N = N-(I*20);
	R = N;
	printf("%d nota(s) de R$ 20.00\n", (int)I);
	I = (int)N/10;
	N = N-(I*10);
	R = N;
	printf("%d nota(s) de R$ 10.00\n", (int)I);
	I = (int)N/5;
	N = N-(I*5);
	R = N;
	printf("%d nota(s) de R$ 5.00\n", (int)I);
	I = (int)N/2;
	N = N-(I*2);
	R = N;
	printf("%d nota(s) de R$ 2.00\n", (int)I);
	
	
	printf("MOEDAS:\n");
	I = (int)N/1;
	N = N-(I*1);
	R = N;
	printf("%d moeda(s) de R$ 1.00\n", (int)I);
	I = N/0.50;
	N = N-((int)I*0.50);
	R = N;
	printf("%d moeda(s) de R$ 0.50\n", (int)I);
	I = N/0.25;
	N = N-((int)I*0.25);
	R = N;
	printf("%d moeda(s) de R$ 0.25\n", (int)I);
	I = N/0.10;
	N = N-((int)I*0.10);
	R = N;
	printf("%d moeda(s) de R$ 0.10\n", (int)I);
	I = N/0.05;
	N = N-((int)I*0.05);
	R = N;
	printf("%d moeda(s) de R$ 0.05\n", (int)I);	
	I = N/0.01;
	N = N-((int)I*0.01);
	R = N;
	printf("%d moeda(s) de R$ 0.01\n", (int)I);	
	

	getch();
	return 0;
}