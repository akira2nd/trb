#include <stdio.h>
#include <conio.h>

int troca (int a,int b);          //definição de função

main ( ){
 int num1,num2;                     //definição de variáveis
 num1=100;                          //variável num1 recebe 100
 num2=200;                          //variável numa2 recebe 200
 troca (num1,num2);               //chama função troca passando endereço de memória
 printf ("\nEles agora valem %d %d\n",num1,num2);//saída “Eles agora valem 200 100”
 getch();
}
int troca (int a,int b){
 int temp;                          //variável temp do tipo inteiro
 temp=a;                           //temp recebe o valor do endereço de memória a(100)
 a=b;                             //a recebe valor do endereço de memória de b(200)
 b=temp;                           //b recebe valor de temp(100)
}
