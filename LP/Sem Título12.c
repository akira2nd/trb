#include <stdio.h>
#include <stdlib.h>

main(){
	int c,n;

	printf("Digite o valor de n:");
	scanf("%d",&n);
	printf("Os %d primeiros impares sao:\n",n);
	c = 1;
	while(c <= n){
        if(c%2 > 0)
		{
			printf("%d\n", c);
		}
		c++;
	}
	getch();
}