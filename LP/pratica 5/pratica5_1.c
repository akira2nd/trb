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

void addCarros(){
	int i = (20 - (sizeof(vetorCarro)/sizeof(carro)));
	gets(vetorCarro[i].marca);
	scanf("%d", &vetorCarro[i].ano);
	gets(vetorCarro[i].cor);
	scanf("%lf", &vetorCarro[i].preco);
}

double ipreco(double x){
	int i;
	for (i = 0; i < (sizeof(vetorCarro)/sizeof(carro); i++)
	{
		if (vetorCarro[i].preco <= x)
		{
			printf("%s\n", vetorCarro[i].marca);
			printf("%s\n", vetorCarro[i].cor);
			printf("%d\n", vetorCarro[i].ano);
		}
	}
}



main()
{
	
}