#include <stdio.h>
#include <conio.h>

main(){

	int a,b;
	scanf("%d %d", &a,&b);

	if (a>=b)
	{
		printf("O JOGO DUROU %d HORA(S)\n", (b+24)-a);
		return 0;
	}
	printf("O JOGO DUROU %d HORA(S)\n", b-a);
	return 0;
}