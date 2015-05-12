#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

int potencia(int x, int z);

main()
{
	int x,z,res;
	printf("Digite um numero:");
	scanf("%d",&x);
	printf("Digite uma potencia:");
	scanf("%d",&z);
	res = potencia(x,z);
	printf("%d elevado a %d = %d\n", x,z,res);
	getch();
}

int potencia(int x, int z)
{
	if(z>=1)
	{
		z--;
		return x * potencia(x , z);
	}
	else
	{
		return 1;
	}
}
/* 6. Escreva uma função recursiva que receba, por parâmetro, dois valores X e Z e calcula e 
	retorna X^Z. (sem utilizar funções prontas) 

OK
*/