#include <stdio.h>
#include <stdlib.h>

main(){
	int n;

	printf("\tCalculo dos quadrados de uma sequencia de numeros\n\n");
	printf("Entre com uma sequencia de numeros inteiros não-nulos, seguidas por 0:\n");
	do{
	scanf("%d",&n);
	printf("O quadrado do numero %d é %d\n", n,(n*n));
	}while(n>0);
	getch();
}