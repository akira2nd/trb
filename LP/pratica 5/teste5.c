#include <stdio.h>
#include <conio.h>
#include <string.h>

struct carro
{
	char marca[15];
	int ano;
	char cor[10];
	double preco;
};

struct carro vetorCarro[20];

main()
{
	printf("MARCA\n");
	gets(vetorCarro[0].marca);
	printf("ANO\n");
	scanf("%d", &vetorCarro[0].ano);
	printf("COR\n");
	gets(vetorCarro[0].cor);
	printf("PRECO\n");
	scanf("%lf", &vetorCarro[0].preco);

	printf("MARCA\n");
	gets(vetorCarro[0].marca);
	printf("ANO\n");
	scanf("%d", &vetorCarro[0].ano);
	printf("COR\n");
	gets(vetorCarro[0].cor);
	printf("PRECO\n");
	scanf("%lf", &vetorCarro[0].preco);

	printf("%d\n", (sizeof(vetorCarro)/sizeof(carro)));
	getch();
}