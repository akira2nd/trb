#include <stdio.h>
#include <conio.h>
#include <complex>

int main()
{
	int A,B,C,R;
	scanf("%d %d %d", &A,&B,&C);
	R = (A+B+ abs(A-B))/2;
	printf("%d eh o maior\n", (R+C+abs(R-C))/2);
    return 0;
}