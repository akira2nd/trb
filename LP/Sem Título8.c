#include <stdio.h>
#include <stdlib.h>

main(){
       int n,c;
       
       printf("Digite um num:");
       scanf("%d",&n);

       for(c=0;c<=12;c++){
                          printf("%d x %d\t= %d\n", n,c,(n*c));
                          }
       getch();
       }
