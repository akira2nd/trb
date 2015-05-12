#include <stdio.h>
#include <conio.h>

main()
{
	int i,a=0;
	int X = 60, Y = 90, V=0;
	scanf("%d",&V);
	
	for (i = 0;a!=V; ++i)
	{
		a = ((Y*i)-(X*i))/60;
	}
	printf("%d minutos\n", i-1);

	return 0;
}