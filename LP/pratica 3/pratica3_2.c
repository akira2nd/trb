#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main()
{
	int vetor[8],vetorz[8];
	int x=0,z=4;
	
	do
	{
	printf("Digite um numero:");
	scanf("%d", &vetor[x]);
	x++;
	}
	while(x<8);

	for(x=0;x<8;x++)
	{
		if (z==8){z=0;}
		vetorz[x] = vetor[z];
		z++;
	}
	for(x=0;x<8;x++)
	{
		printf("%d", vetorz[x]);
	}
	
	getch();
}
/*2. Leia um vetor de 8 posições e troque os 4 primeiros valores
pelos 4 últimos e vice e versa. Escreva ao final o vetor obtido.*/