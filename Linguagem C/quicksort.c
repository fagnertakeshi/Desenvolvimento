#include<stdio.h>

#define NUMAX 10000
#define MAX 10000

void vetorAleatorio(int v[], int);
void imprimeVetor(int v[], int n);
void qsort(int v[],int left, int right);
void swap (int v[],int i, int j);

int main()
{
int v[NUMAX];

vetorAleatorio(v,MAX);
qsort(v,1,NUMAX);
//imprimeVetor(v, NUMAX);
	
}

void vetorAleatorio(int v[], int n)
{
int i;
	
	srand(2);
	for (i=0; i<n; ++i)
		v[i]=rand()%n;	
}
void imprimeVetor(int v[], int n)
{
int i;

	for (i=0; i<n ;++i)
		printf("%d  ",v[i]);

}


void qsort(int v[],int left, int right)
{
	int i,last;
	void swap(int v[],int i, int j);
	
	if (left >= right) 
		return;
	swap(v,left,(left+right)/2);
	last=left;
	for (i=left+1;i <=right;i++)
		if (v[i]<v[left])
			swap(v,++last,i);
	swap(v,left,last);
	qsort(v,left,last);
	qsort(v,last+1,right);
}

void swap (int v[],int i, int j)
{
	int temp;
	temp =v[i];
	v[i]=v[j];
	v[j]=temp;
}
