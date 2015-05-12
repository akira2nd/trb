#include <stdio.h>
#include <conio.h>


main()
{
	double n1,n2,n3,n4,media,ne;
	
	scanf("%lf %lf %lf %lf", &n1,&n2,&n3,&n4);
	n1 *= 2;
	n2 *= 3;
	n3 *= 4;
	n4 *= 1;
	media = (n1+n2+n3+n4)/10;
	printf("Media: %.1lf\n", media);

	if (media >= 7.0)
	{
		printf("Aluno aprovado.\n");
		getch();
		return 0;
	}else
	{
		if (media <5.0)
		{
			printf("Aluno reprovado.\n");
			getch();
			return 0;
		}else
		{
			printf("Aluno em exame.\n");
			scanf("%lf", &ne);
			printf("Nota do exame: %.1lf\n", ne);
			media = (media + ne)/2;
			if (media>5.0)
			{
				printf("Aluno aprovado.\n");
				printf("Media final: %.1lf\n", media);
				getch();
				return 0;
			}else
			{
				printf("Aluno aprovado.\n");
				printf("Media final: %.1lf\n", media);
				getch();
				return 0;
			}
		}
	}
	getch();
	return 0;
}