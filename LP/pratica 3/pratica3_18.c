#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main(){
	int vetor[2][3][2];
	int a=1,l,c,i;

	for (l=0;l<2;l++)
	{
		for (c = 0; c < 3; c++)
		{
			for (i = 0; i < 2; i++)
			{
				printf("Elemento[%d][%d][%d] = ", l,c,i);
				scanf("%d", &vetor[l][c][i]);
			}
		}
	}

	printf("\n\n***********Saida de dados***********\n\n");
	for (l=0;l<2;l++)
	{
		for (c = 0; c < 3; c++)
		{
			for (i = 0; i < 2; i++)
			{
				//vetor[l][c][i] = a;
				//a++;
				printf("Elemento[%d][%d][%d] = %d\n", l,c,i,(vetor[l][c][i]));
			}
		}
	}
	getch();
}
/*18.Carregar e imprimir uma matriz[2][3][2].*/