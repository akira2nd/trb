#include <stdio.h>
#include <conio.h>

main(){
	
	int a,b;
	scanf("%d %d", &a,&b);

	if (!(a%b))
	{
		printf("Sao Multiplos\n");
		return 0;
	}else
	{
		if (!(b%a))
		{
			printf("Sao Multiplos\n");
			return 0;
		}else
		{
			printf("Nao sao Multiplos\n");
			return 0;
		}
	}
	return 0;
}
