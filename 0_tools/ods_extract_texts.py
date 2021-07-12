#!/usr/bin/python3
# -*- coding:utf8 -*-

import pandas as pd 
from sys import argv
from os import mkdir
from os.path import exists, join

def folder_create(path):
	if not exists(path):
		mkdir(path)

def main(file):
	doc = pd.ExcelFile(file)
	folder_export = './texts/'
	s = ' '*4
	
	folder_create(folder_export)

	for page in doc.sheet_names:
		print(page)
		file_export = join(folder_export, f'{page}.yml')
		df = doc.parse(page)
		df.fillna('', inplace=True)

		dic = {}
		for i in range(len(df)):
		    par = {}
		    for c in df.columns:
		        par[c] = df[c][i]

		    dic[i] = par

		text = []
		text.append(f'title: "{page}"')
		text.append(f'\nparagraphs:')

		for k in dic.keys():
			text.append(f'\n{s*1}-')
			for i in dic[k].keys():
				text.append(f'{s*2}{i}: |\n{s*3}{dic[k][i]}')

		with open(file_export, 'w') as f:
			f.write('\n'.join(text))

if __name__ == '__main__':
	main(argv[1])