#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

int snprimos(int n);

main()
{
	int n,soma;
	printf("Numero:");
	scanf("%d", &n);
	soma = snprimos(n);
	printf("%d",soma);
	getch();
}

int snprimos(int n)
{
	int d=2,num=2,soma=0,cont=1;
	while(n!=0)
	{
		if(d != num && not(num%d))
		{
			d = 2;
			num++;
		}
		else if (d==num)
		{
			printf("%d", num);
			soma += num; 
			num ++;
			d = 2;
			n--;
			if(n!=0){printf(" + ");}
			else{printf(" = ");}
		}
		else if (num%d){d++;}
	}
	return soma;
}
//2. Escreva uma função que leia um inteiro não-negativo n e imprima a soma dos n primeiros números primos.