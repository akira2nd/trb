#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main()
{
	int vetorA[5] = {5,8,7,3,4}, vetorB[8] = {11,2,8,9,10,1,5};
	int x,i;

	for (i = 0; i < (sizeof(vetorA)/sizeof(int)); i++)
	{
		for (x = 0; x < (sizeof(vetorB)/sizeof(int)); x++)
		{
			if (vetorA[i] == vetorB[x])
				printf("%d,", vetorA[i]);
		}
	}

	//printf("%d", (sizeof(vetorA)/sizeof(int)));
	getch();
}

/*4. Dado dois vetores, A (5 elementos) e B (8 elementos), faça um programa em C que
imprima todos os elementos comuns aos dois vetores.
Exemplo: int A[5] = {1,2,4,6,21};
		int B[8] = {2,3,6,7,9,11,15,20};
*/