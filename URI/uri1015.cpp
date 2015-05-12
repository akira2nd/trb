#include <stdio.h>
#include <conio.h>
#include <math.h>

main()
{
	double aX,bX,aY,bY,res;
	scanf("%lf %lf", &aX,&aY);
	scanf("%lf %lf", &bX,&bY);
	res = sqrt(((bX - aX)*(bX - aX)) + ((bY - aY)*(bY - aY)));
	printf("%.4lf\n", res);
	getch();
	return 0;
}