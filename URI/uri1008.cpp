#include <stdio.h>
#include <conio.h>

int main()
{
	int numero, hora;
	double salarioh;

	scanf("%d %d %lf", &numero,&hora,&salarioh);
	salarioh *= hora;
	printf("NUMBER = %d\nSALARY = U$ %.2lf\n", numero,salarioh);
    return 0;
}