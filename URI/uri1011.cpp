#include <stdio.h>
#include <conio.h>

int main()
{
	long int numero;
	double VOLUME;

	scanf("%d", &numero);
	printf("VOLUME = %.3lf\n", ((4.0/3) * 3.14159 * numero*numero*numero));
	getch();
    return 0;
}