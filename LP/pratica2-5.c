#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

int fatorial(int x);

main()
{
	int n, res;

	printf("Fatorial\n\nDigite um numero:");
	scanf("%d", &n);
	res = fatorial(n);
	printf("O fatorial de %d é %d", n,res);
	getch();
}

int fatorial(int x)
{
	if (x != 0)
	{
		return x * fatorial(x-1);
	}
	else
	{
		return 1;
	}
}

/* 5. Faça uma função para calcular o fatorial de um número fornecido pelo usuário.
    A função fatorial de um número natural n é o produto de todos os n primeiros números naturais.
    Fat(n)=n!=1.2.3.4...n. Vamos tomar Fat(0)=1.
     OBS: Resolver de forma recursiva. (Utilizar pesquisa da parte B)

OK
*/
