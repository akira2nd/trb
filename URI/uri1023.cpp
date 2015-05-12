#include <stdio.h>
#include <conio.h>
#include <math.h>

main()
{
	double A,B,C,delta,R1,R2;
	scanf("%lf %lf %lf", &A,&B,&C);
	delta = (B*B)-(4*A*C);
	if (delta<0)
	{
		printf("Impossivel calcular\n");
		return 0;
	}
	R1 = -(B)+sqrt(delta);
	R2 = -(B)-sqrt(delta);
	if (!R1 || !R2)
	{
		printf("Impossivel calcular\n");
		return 0;

	}
	printf("R1 = %.5lf\n", (R1/(2*A)));
	printf("R2 = %.5lf\n", (R2/(2*A)));
	getch();
	return 0;
}