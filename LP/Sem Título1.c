#include <stdio.h>
#include <stdlib.h>

main()
{
	int a,b,soma;
	printf("Digite o primeiro n�mero:\n");
	scanf("%d",&a);
	printf("Digite o segundo n�mero:\n");
	scanf("%d",&b);
	soma = a + b;
	printf("A soma dos n�meros %d + %d e igual %d\n", a,b,soma);
	getch();
}
