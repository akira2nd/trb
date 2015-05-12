#include <stdio.h>
#include <conio.h>
#include <string.h>
#include <stdlib.h>

//variaveis e struct global
int _i = 0;
struct carros
{
	char marca[15];
	int ano;
	char cor[10];
	double preco;
};
struct carros vetcarros[20];

//inicio das funções
//função para cadastrar carros, marca ano cor preço
int ler_carros(int i){

	printf("Digite a marca:\n");
	scanf("%s", &vetcarros[i].marca);

	printf("\nDigite o ano:\n");
	scanf("%d", &vetcarros[i].ano);

	printf("\nDigite a cor:\n");
	scanf("%s", &vetcarros[i].cor);

	printf("\nDigite o preco:\n");
	scanf("%lf", &vetcarros[i].preco);

	printf("\n");
	system("PAUSE");
	_i++;
}

//função para procurar por preço
double ler_preco(double x, int i){
	int a;
	
	printf("\nLocalizado(s):\n");
	for (a = 0; a < i; a++)
	{
		if (vetcarros[a].preco <= x)
		{
			printf("marca:\t%s\n", vetcarros[a].marca);
			printf("cor:\t%s\n", vetcarros[a].cor);
			printf("ano:\t%d\n", vetcarros[a].ano);
			printf("\n");
		}
	}
	printf("\n");
	system("PAUSE");
}

//função para procurar por marca
char ler_marca(char x[15], int i){
	int a;
	
	printf("\nLocalizado(s):\n");
	for (a = 0; a < i; a++)
	{
		if (!stricmp(vetcarros[a].marca,x))
		{
			printf("preco:\t%.2lf\n", vetcarros[a].preco);
			printf("ano:\t%d\n", vetcarros[a].ano);
			printf("cor:\t%s\n", vetcarros[a].cor);
			printf("\n");
		}
	}
	printf("\n");
	system("PAUSE");
}

//função para procurar por marca ano cor
char ler_m_a_c(char m[15], int a, char c[10], int b){
	int i,nada = 0;

	printf("\nLocalizado(s):\n");
	for (i = 0; i < b; i++)
	{
		if (!stricmp(vetcarros[i].marca,m) && vetcarros[i].ano == a && !stricmp(vetcarros[i].cor, c))
		{
			printf("preco: %.2lf\n", vetcarros[i].preco);
			printf("\n");
			nada ++;
		}
	}
	if (!nada)
	{
		printf("Nenhum registro com essas especificacoes\n");
	}
	printf("\n");
	system("PAUSE");
}

//função para chamar as demais....................................................................................................
int escolher_func(int x, int i){
	double pr;
	char m[15], c[10];
	int a;
	system("cls");

	switch(x){
		case 1:
		ler_carros(i);
		break;
		
		case 2:
		printf("Digite o preco procurado\n");
		scanf("%lf", &pr);
		ler_preco(pr, i);
		break;
		
		case 3:
		printf("Digite a marca procurada\n");
		scanf("%s", &m);
		ler_marca(m, i);
		break;

		case 4:
		printf("Digite a marca, ano e cor desejado:\n");
		scanf("%s %d %s", &m, &a, &c);
		ler_m_a_c(m, a,c,i);
		break;

	}
}

//inicio do programa.........................................................................................................
main()
{
	int x;

	do
	{
		printf("\t**** Carros ****\n\n");
		printf("( 1 )---Cadastrar Carro\n");
		printf("( 2 )---Procurar por preco\n");
		printf("( 3 )---Procurar por marca\n");
		printf("( 4 )---Procurar por marca, ano e cor\n");
		printf("( 5 )---Sair\n");

		printf("Selecione uma atividade:\n");
		
		scanf("%d", &x);
		if (x <=5 && x!=0)
		{
			escolher_func(x, _i);
		}else{
			system("cls");
			printf("\nOpcao invalida\n\n");
			system("PAUSE");
		}
		system("cls");
		
	} while (x != 5);
}