#include <stdio.h>
#include <conio.h>

int main()
{
	double A,B,C;
	scanf("%lf %lf %lf", &A, &B, &C);
	printf("TRIANGULO: %.3lf\n", (A*C/2));
	printf("CIRCULO: %.3lf\n", (3.14159*C*C));
	printf("TRAPEZIO: %.3lf\n", (C*(A+B)/2));
	printf("QUADRADO: %.3lf\n", (B*B));
	printf("RETANGULO: %.3lf\n", (A*B));
    return 0;
}
/*
Escreva um programa que leia três valores com ponto flutuante de dupla precisão: A, B e C. Em seguida, calcule e mostre: 
a) a área do triângulo retângulo que tem A por base e C por altura.	A * C / 2
b) a área do círculo de raio C. (pi = 3.14159)						pi = 3.14159 * C^2
c) a área do trapézio que tem A e B por bases e C por altura.		C*(A+B) / 2
d) a área do quadrado que tem lado B. 								B * B
e) a área do retângulo que tem lados A e B. 						A * B
*/