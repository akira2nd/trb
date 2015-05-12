#include <stdio.h>
#include <conio.h>


main()
{
	int cod,item;
	double vetor[] = {4.00,4.50,5.00,2.00,1.50};
	
	scanf("%d %d", &cod,&item);
	printf("Total: R$ %.2lf\n", (vetor[cod-1]*item));

	getch();
	return 0;
}