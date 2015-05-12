#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main()
{
	int vetor[10] = {0,2,42,43,6,7,8,10,11,15};
	int i,n;

	printf("Digite um numero para encontrar no vetor: ");
	scanf("%d",&n);
	for (i = 0; i < 10; i++)
	{
		if(vetor[i]== n)
		{
			printf("O numero foi encontrado na posicao [%d] do vetor", i);
			break;
		}else if (i == 9)
		{
			printf("O numero nao foi encontrado na vetor", i);
			break;
		}
	}
	getch();
}
/*3. Inicialize um vetor de 10 posições e em seguida leia um valor X qualquer.
Seu programa devera fazer uma busca do valor de X no vetor lido 
e informar a posição em que foi encontrado ou se não foi encontrado.
Exemplo: int vetor[10] = {2,5,4,54,43,22,5,9,30,15};*/