#include <stdio.h>
#include <stdlib.h>

main(){
       int n, a,fat;
       printf("Digite um numero:");
       scanf("%d",&n);
       fat = 1;
       for(a=1 ; a<n ; a++){
                         fat = fat + fat*a;
                         }
       printf("O fatorial de %d é: %d", n,fat);
       getch();
}
