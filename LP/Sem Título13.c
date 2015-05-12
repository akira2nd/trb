#include <stdio.h>
#include <stdlib.h>

main()
{
	int c,x,n,res;
	printf("\tCalculo de potencias\n\nDigite um numero inteiro:");
	scanf("%d",&x);
	printf("Digite um numero inteiro nao negativo:");
	scanf("%d",&n);
	res = x;
	for(c=1;c<n;c++){
		res = res * x;
	}
	printf("\nO valor de %d elevado a %d = %d", x,n,res);
	getch();

}
