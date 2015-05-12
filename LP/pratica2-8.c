#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

int epar(int x);

main()
{
	int x;
	printf("Digite um numero:");
	scanf("%d",&x);
	if(epar(x))
	{
		printf("O número %d é par!\n", x);
	}
	else
	{
		printf("O número %d é impar\n", x);
	}
	getch();
}

int epar(int x)
{
	return not(x%2);
}
/* 8.	Escreva uma função que recebe um inteiro positivo n e devolve 1 se n é par e 0 se n é impar. */