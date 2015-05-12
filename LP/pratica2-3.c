#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

int divisor(int cont);

main()
{
	int n,num;
	printf("Numero:");
	scanf("%d",&n);
	num = divisor(n);
	printf("%d possui %d divisores\n", n,num);
	getch();
	//system ("pause");
}

int divisor(int cont)
{
	int i,n=0;
	for (i = cont; i != 0; i--)
	{
		if(cont%i == 0)
		{
			n++;
		}
	}
	return n;
}

/*3. Faça uma função que receba, por parâmetro, um valor inteiro e positivo e 
	retorna o número de divisores desse valor.

OK
*/