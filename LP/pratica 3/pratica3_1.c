#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main()
{
	int count=0,n;
	int vetor[10000];
	
	printf("Digite a quantidade de numeros a serem digitados:");
	scanf("%d", &n);
	
	do
	{
		printf("Digite um numero: ");
		scanf("%d",&vetor[count]);
		count++;
	}while(count<n);
	
	printf("Sequencia na ordem inversa\n");
	for(--n;n>=0;n--)
	{
		printf("%d ", vetor[n]);
	}
	getch();
}
/* 1. Dada uma sequência de n números, imprimi-la na ordem inversa a da leitura. */