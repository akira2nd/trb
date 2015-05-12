#include <stdio.h>
#include <conio.h>

int main()
{
	char nome[20];
	double salario, valorvds;

	scanf("%s %lf %lf", &nome, &salario, &valorvds);
	printf("TOTAL = R$ %.2lf\n", salario+(valorvds*0.15));
    return 0;
}