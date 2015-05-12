#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main(){
	int vetor[10][10];
	int l,c,maior, lin,col;

	maior = vetor[0][0];
	for (l = 0; l < 10; l++)
		{
			for (c = 0; c < 10; c++)
			{
				printf("vetor[%d][%d] = %d\n", l,c,(vetor[l][c]));
				if (maior<vetor[l][c]){
					maior = vetor[l][c];
					lin = l;
					col = c;
				}
			}
		}	
	printf("O vetor de maior valor e o vetor[%d][%d]", lin,col);
	getch();
}
/*7. Leia uma matriz 10 x 10 e escreva a localização (linha e a coluna) do maior valor.*/