#include <stdio.h>
#include <stdlib.h>

main()
{
	int c,n,soma;

	printf("\tCalculo da soma dos n primeiros números naturais\n\nDigite o valor de n: ");
	scanf("%d",&n);
	soma = n;
	for (c = n-1; c > 0; c--)
	{
		soma = soma + c;
	}

	printf("A soma dos %d primeiros numeros naturais é %d", n,soma);
	getch();
}