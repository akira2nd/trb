#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

int triangulo(int x, int y, int z);

main()
{
	int x, y, z;

	printf("Digite 3 medidas de lados:\n1º lado:");
	scanf("%d",&x);
	printf("2º lado:");
	scanf("%d",&y);
	printf("3º lado:");
	scanf("%d",&z);
	triangulo(x,y,z);
	getch();
}

int triangulo(int x, int y, int z)
{
	if(x==0 || y ==0 || z==0)
	{
		return printf("Não é triangulo");
	}
	else if(x == y && x == z)
	{
		return printf("O Triangulo é Equilatero!");
	}
	else if(x != y && x != z && z != y)
	{
		return printf("O Triangulo é Escaleno!");
	}
	else
	{
		return printf("O Triangulo é Isoceles!");
	}
}

/* 

	4. Escreva uma função que recebes 3 valores inteiros e positvos 
	X, Y e Z e que verifique se esses valores podem ser os comprimentos dos lados de um triângulo e, 
	neste caso, retornar qual o tipo de triângulo formado. Para que X, Y e Z 
	formem um triângulo é necessário que a seguinte propriedade seja satisfeita: 
	o comprimento de cada lado de um triângulo é menor do que a soma do comprimento dos outros dois lados. 
	O procedimento deve identificar o tipo de triângulo formado observando as seguintes definições:
•	o Triângulo Equilátero: os comprimentos dos 3 lados são iguais.
•	o Triângulo Isósceles: os comprimentos de 2 lados são iguais.
•	o Triângulo Escaleno: os comprimentos dos 3 lados são diferentes.

OK
*/