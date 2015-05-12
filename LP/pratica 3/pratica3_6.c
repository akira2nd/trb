#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main()
{
	int linha = 5, coluna = 5;
	int vetor[linha][coluna];
	int c,l,a=1;

	for (l = 0; l < 5; l++)
	{
		for (c = 0; c < 5; c++)
		{
			if (l == c)
				vetor[l][c] = 1;
			else
				vetor[l][c] = 0;
		}
	}

	for (l = 0; l < 5; l++)
	{
		for (c = 0; c < 5; c++)
		{
			printf("%d", vetor[l][c]);
		}
		printf("\n");
	}
	getch();
}
/* 6. Declare uma matriz 5 x 5. 
Preencha com 1 a diagonal principal e com 0 os demais elementos.
Escreva ao final a matriz obtida.
*/