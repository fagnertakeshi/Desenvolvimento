#include <stdio.h>
#include <stdlib.h>
#define NUMMAX 200
#define MAX 100000
void shellsort(int v[],int n);
void bublesort(int v[],int n);
main() 
{
	int i;
	int srand[MAX];
	
	for (i=0;i<MAX;i++)
		srand[i]= rand() % NUMMAX;	
		
	shellsort(srand,MAX);
	//bublesort(srand,MAX);
	
	
}

void shellsort(int v[],int n){
	int gap,i,j,temp;
	
	for (gap=n/2;gap>0;gap/=2)
		for (i=gap;i<n;i++)
			for (j=i-gap;j>=0 && v[j] > v[j+gap];j-=gap) {
				temp=v[j];
				v[j]=v[j+gap];
				v[j+gap]=temp;
			}
			
	
			
}

void bublesort(int v[],int n){
	int ntrocas,i,j,temp;
	ntrocas=1;
	while (ntrocas){
	ntrocas=0;	
	for (i=0;i<n-1;i++)
		if (v[i]>v[i+1]) {
			ntrocas++;
			temp=v[i];
			v[i]=v[i+1];
			v[i+1]=temp;
		}
	}
	/*for (i=0;i<n;i++)
		printf("%6d%c",v[i],(i%10==9||i==n-1)?'\n':' ');*/	
}
