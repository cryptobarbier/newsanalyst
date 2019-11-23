# -*- coding: utf-8 -*-
"""
Created on Sat Nov 23 19:34:15 2019

@author: PC
"""
import requests
import json
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import linear_kernel
from sklearn.cluster import AgglomerativeClustering
from summarizer import Summarizer
from docx import Document
from docx.shared import Pt




#Extract the site rank of json result
def Rank(sce):
  try:
    return sce['ranking']['alexaGlobalRank']
  except:
    return 1500000

# Extract Number of Facebook shares
def Shares(sh):
  try:
    return sh['facebook']
  except:
    return 0

class NewsArticle():
    def __init__(self,req,cutoff_ranking=30000,weight_cutoff=20):
            #Run the request and put the results into a json file
        res=requests.get(req)
        j=json.loads(res.text[14:-1],encoding='ascii')['articles']['results']
        df = pd.DataFrame.from_dict(j, orient='columns')
        #Remove irrelevant articles
        df=df[df['wgt']>weight_cutoff]
        #Remove Duplicates
        df=df[df['isDuplicate']==False]
        # Remove duplicate bodys i.e articles published in several websites
        df.drop_duplicates(subset ='title', 
                             keep = 'first', inplace = True) 
        # Filter the articles on source rating
        df['Source Rank']=df['source'].apply(Rank)
        df['Shares']=df['shares'].apply(Shares)
        # replace the /n
        df['body']=df['body'].str.replace('\n',' ')
        df['body']=df['body'].str.replace("\'s","'s")
        df.sort_values(by='Source Rank')
        df=df[df['Source Rank']<cutoff_ranking]
        self.results=df
    
    def CreateDist(self):
        
        # Build the text database
        corp=[self.results.loc[x]['body'] for x in self.results.index ]
        tfidf = TfidfVectorizer().fit_transform(corp)
        cosine_similarities = linear_kernel(tfidf, tfidf)
        matrix=pd.DataFrame(cosine_similarities,index=self.results.index,columns=self.results.index)
        self.dist=1-matrix
        
    def Cluster(self,thresh=0.4):
        clustering = AgglomerativeClustering(affinity='precomputed',distance_threshold=thresh,linkage='complete',n_clusters=None).fit(self.dist)
        self.clust=pd.DataFrame(clustering.labels_,index=self.results.index).sort_values(by=0)
    
    def Summary(self):
        summ=dict.fromkeys(self.clust[0].unique())
        # Will store the name of Cluster based on article name
        subtitles=[]
        # Will store the url of main article summarizing the cluster
        urls=[]
        model = Summarizer()
        # Create the summary by aggregating all the bodys of the articles
        for c in summ.keys():
            df2=self.clust[self.clust[0]==c]
            subtitles=subtitles+[self.results.loc[df2.index[0]]['title']]
            urls=urls+[self.results.loc[df2.index[0]]['url']]
            full_text=' '.join(self.results.loc[df2.index]['body'].unique())
            full_text=full_text.replace('\n',' ')
            length=len(full_text)
            if length<500:
              target=200
              ratiol=min(1,target/length)
            else:
              target=min(1000,length*0.2)
              ratiol=target/length
            summ[c]=model(full_text,min_length=60,ratio=ratiol)
        self.summary=summ
        self.subtitles=subtitles
        self.urls=urls
        
    def TextOutput(self,title,docname):
          document = Document()
          style = document.styles['Normal']
          font = style.font
          font.name = 'Times New Roman'
          font.size = Pt(10)
          # Write the title passed at the function
          document.add_heading(title, 0)
          i=0
          for element in self.summary:
            document.add_heading(self.subtitles[i], level=1)
            u=document.add_heading(self.urls[i], level=2)
            u.italic=True
            u.style = document.styles['Normal']
            p = document.add_paragraph(self.summary[element])
            i=i+1
          document.save(docname+'.docx')
          self.document=document
    
    
        

    
            