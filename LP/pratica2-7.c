#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

int ler();
int multiplicar(int x, int y);
int imprimir(int x);

main()
{
	imprimir(ler());
	getch();
}

int ler()
{
	int x,y;
	printf("Digite o primeiro numero:");
	scanf("%d",&x);
	printf("Digite o segundo numero:");
	scanf("%d",&y);
	return multiplicar(x,y);
}

int multiplicar(int x, int y)
{
	return x * y;
}

int imprimir(int x)
{
	printf("%d\n",x);
}

/* 7. Faça 3 funções:
•	ler: - uma função que receba dois número inteiro positivo n;
•	multiplicar: - que multiplique os números recebidos na função ler;
•	imprimir: - que imprima o resultado da função multiplicar.
*/