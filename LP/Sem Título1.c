#include <stdio.h>
#include <stdlib.h>

main()
{
	int a,b,soma;
	printf("Digite o primeiro número:\n");
	scanf("%d",&a);
	printf("Digite o segundo número:\n");
	scanf("%d",&b);
	soma = a + b;
	printf("A soma dos números %d + %d e igual %d\n", a,b,soma);
	getch();
}
