#include <stdio.h>
#include <stdlib.h>

main(){
       int a, n1,n2,res;
       
       do{
       printf("Digite a operação desejada:\n");
       printf("\t( 1 )...Divisao\n\t( 2 )...Resto da divisao\n\t( 3 )...Adição\n\t( 4 )...Multiplicação\n");
       
       scanf("%d",&a);
       
       }while(a > 4);
       
       printf("Entre com o primeiro numero:");
       scanf("%d",&n1);
       printf("Entre com o segundo numero:");
       scanf("%d",&n2);
      
       switch(a){
                 case 1:
                      res = n1/n2;
                      break;
                 case 2:
                      res = n1%n2;
                      break;
                 case 3:
                      res = n1 + n2;
                      break;
                 case 4:
                      res = n1 * n2;
                      break;
                 }
       printf("Resultado: %d", res);
       
       getch();
}
