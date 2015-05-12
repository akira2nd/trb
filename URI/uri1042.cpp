#include <stdio.h>
#include <conio.h>


main()
{
	int vetor[3],i,f,guardar,vetorc[3];
	scanf("%d %d %d", &vetor[0],&vetor[1],&vetor[2]);
	
	for (i = 0; i < 3; i++)
	{
		vetorc[i] = vetor[i];
	}

	for (i = 0; i < 3; i++)
	{
		for (f = 0; f < 3; f++)
		{
			if (vetor[i]<vetor[f])
			{
				guardar = vetor[f];
				vetor[f] = vetor[i];
				vetor[i] = guardar;
			}
		}
	}

	for (i = 0; i < 3; i++)
	{
		printf("%d\n", vetor[i]);
	}
	
	printf("\n");
	for (i = 0; i < 3; i++)
	{
		printf("%d\n", vetorc[i]);
	}

	getch();
	return 0;
}