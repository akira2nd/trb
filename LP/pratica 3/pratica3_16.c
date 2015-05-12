#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

main(){
	double vetor[500];
	int cont=0, i;

	printf("Digite uma sequencia de numeros: -999 para parar\n");
	do{
		scanf("%d", &vetor[cont]);
		printf("%d\n", vetor[cont]);
	}while(vetor[cont] != (double)-999);
	getch();
}

/* 16. Dada uma sequência de n números reais, determinar o número de vezes que cada um deles ocorre na mesma.
Exemplo: n = 8
Seqüência: -1.7, 3.0, 0.0, 1.5, 0.0, -1.7, 2.3, -1.7
Saída: -1.7 ocorre 3 vezes
   	    3.0 ocorre 1 vez
       	0.0 ocorre 2 vezes
        1.5 ocorre 1 vez
        2.3 ocorre 1 vez
*/
