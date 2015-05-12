#include <stdio.h>
#include <stdlib.h>
#include <conio.h>
#include <string.h>

main(){
	char Texto[51];
	char Texto1[51];
	char a = 0, count = 0, i;

	printf("Digite uma frase até 50 caracteres:\n");
	gets(Texto);
	for (i = 0; i < strlen(Texto); i++)
	{
		if (Texto[i] == ' ')
		{
			count++;
		}else{
			Texto1[a] = Texto[i];
			a++;
		}
	}
	printf("Texto sem espaço(s): %s\n",Texto1);
	printf("%d espaço(s)\n", count);
	getch();
}

/*10. Fazer um programa em C que leia uma frase de até 50 caracteres (utilizar o comando gets) 
e imprima a  frase sem os espaços em branco. Imprimir também a quantidade de espaços em branco da frase.*/	