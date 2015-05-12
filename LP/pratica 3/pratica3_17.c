#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main(){
	int vetor[2][6];
	int l,c;

	for (l = 0; l < 2; l++)
	{
		printf("\nLinha %d:\n", (l+1));
		for (c = 0; c < 6; c++)
		{
			scanf("%d", &vetor[l][c]);
		}
	}

	for (l = 0; l < 2; l++)
	{
		printf("\n");
		for (c = 0; c < 6; c++)
		{
			printf("%d", (vetor[l][c]));
		}
	}

	getch();
}
/*17. Carregar uma matriz [2] [6] com números inteiros e em seguida imprimi-la.*/