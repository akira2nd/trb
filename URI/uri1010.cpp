#include <stdio.h>
#include <conio.h>

int main()
{
	int p1,np1,p2,np2;
	double	vp1,vp2;

	scanf("%d %d %lf", &p1,&np1,&vp1);
	scanf("%d %d %lf", &p2,&np2,&vp2);

	printf("VALOR A PAGAR: R$ %.2lf\n", ((vp1*np1)+(vp2*np2)));
	getch();
    return 0;
}