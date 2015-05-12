#include <stdio.h>
#include <stdlib.h>
#include <conio.h>
#include <string.h>

char vetorA[9999][30];
int vetorN[9999];

main(){
	int count = 0;
	int teste = 1;
	
	do{
	printf("Digite o nome do aluno:\n");
	scanf("%s", vetorA[count]);
	printf("Digite a nota do Aluno:\n");
	scanf("%d", &vetorN[count]);
	count++;
	printf("Digite 0 para terminar 1 para continuar\n");
	scanf("%d", &teste);
	}while(teste != 0);
	

	printf("%s\n", vetorA);
	printf("%d\n", vetorN[0]);
	getch();
}

/*
8. Fazer um programa em C para ler uma quantidade N de alunos.
Ler a nota de cada um dos N alunos e calcular a média aritmética das notas.
Contar quantos alunos estão com a nota acima de 5.0.
Obs.: Se nenhum aluno tirou nota acima de 5.0, imprimir mensagem:
Não há nenhum aluno com nota acima de 5.*/