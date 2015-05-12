#include <stdio.h>
#include <stdlib.h>

main(){
       int n1,n2,n3;
       int maior, menor, centro;
       printf("Primeiro numero:");
       scanf("%d", &n1);
       printf("Segundo numero:");
       scanf("%d", &n2);
       printf("Terceiro numero:");
       scanf("%d", &n3);     
       if(n1>n2){
                 maior = n1;
                 menor = n2;
                 }
       else{
                 maior = n2;
                 menor = n1;
                 }
       if(n3 > maior){
             centro = maior;
             maior = n3;
                 }
       else if(n3<menor){
            centro = menor;
            menor = n3;
            }
       else{
            centro = n3;
            }
       printf("%d %d %d", menor,centro,maior);       
       getch();
}
