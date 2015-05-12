#include <stdio.h>
#include <conio.h>

main()
{
	int temp, vel;
	scanf("%d %d", &temp, &vel);
	printf("%.3lf\n", (temp*vel)/12.0);
	return 0;
}