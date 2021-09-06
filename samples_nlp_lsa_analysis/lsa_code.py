import sklearn
import pandas as pd
import seaborn as sns
import numpy as np
import scipy as sp
from STAAR_preprocessing_code import *

from sklearn.feature_extraction.text import CountVectorizer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import euclidean_distances
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.manifold import MDS
from mpl_toolkits.mplot3d import Axes3D
from scipy.cluster.hierarchy import ward, dendrogram


#CODE TO ANALYZE DOCUMENT SIMILARITY---------------------------------------------------
#Code to return a document-term-matrix and document indices for a given lesson
#input: path of the lesson folder (for example  'C:/Users/eabalo/Documents/STAAR2014/4g/405B/')
#output: document-term-matrix and document indices as numpy arrays
def dtm_matrix(lessonpath):
    #lesson number
    lessonname = lessonpath.split('/')[-2]

    #creating corpus of txt files
    corpusText(lessonpath)
    
    #finding the paths of the text files
    corpuspath = 'C:/Users/eabalo/Desktop/STAAR35Analyses/data/corpus'
    filepaths = glob.glob(corpuspath + '/'+ lessonname + '/*.txt')

    #script names
    docindex = [w.split('-')[-1].split('.')[0] for w in filepaths]

    #building a document-term matrix
    vectorizer = CountVectorizer(input = 'filename')

    dtm = vectorizer.fit_transform(filepaths)

    #lexicon of words in lesson
    #vocab = vectorizer.get_feature_names()

    #converting to numpy arrays
    dtm = dtm.toarray()

    #vocab = np.array(vocab)
    
    return dtm, docindex, lessonname


#Latent semantic analysis
#Code to return a document-term matrix after applying LSA
#input: dtm (document-term matrix) and ndim (dimension of reduced subspace)
#output: reconstructed dtm from reduced svd (singular value decomposition) matrices
def LSA_dtm(dtm, ndim):
    #SVD
    Umat, Smat, Vmat = np.linalg.svd(dtm)

    #'Re-construction' of document-term matrix using reduced SVD matrices
    Smatd = np.zeros((ndim, ndim), dtype = complex)
    Smatd[:ndim, :ndim] = np.diag(Smat)[:ndim, :ndim]

    Vmatd = Vmat[:ndim,:]
    Umatd = Umat[:,:ndim]

    dtm2 = np.dot(Umatd, np.dot(Smatd, Vmatd))
    
    return dtm2

	
#Code to return indices labeling clone items in a lesson
#input: the lesson's folder path (for example: 'C:/Users/eabalo/Documents/STAAR2014/4g/405B/')
#output: returns a dataframe with lesson item UID and matching index. An index of 1 if the lesson item... 
#...is less than a given distance (0.01 or 5th percentile of the non-zero values in the distance matrix?) of one or more other lesson items. 
#Otherwise it returns an index of 0
def lsa_clone_index(lessonpath):
    #document-term matrix and document indices
    dtm, docindex, lessonname = dtm_matrix(lessonpath)
    
    #reconstructed dtm matrix using LSA and a reduced subspace of dimension 3
    dtm2 = LSA_dtm(dtm, 3)
    
    #distance metric based on cosine similarity
    dist = 1 - cosine_similarity(dtm2)

    dist = np.round(dist, 10)
    
    #threshold distance
    mindist = 0.001
    #mindist = np.percentile(dist[np.nonzero(dist)], 5)

    #set to 1 if less than mindist (or 0 otherwise) to identify items that are clones
    dist[dist > mindist] = 0
    dist[(dist < mindist) & (dist >0)] = 1

    #sum of assigned indices for each lesson item. A number greater than 0 means the item is a clone
    distindex = np.sum(dist, axis = 0)

    #set 1 for any lesson item that is a clone (otherwise 0)
    distindex[distindex > 0] = 1
    distindex = np.round(distindex, 2)
    
    #Changing lesson item index to match UID
    docindex2 = ['TEKS'+ lessonname + '-'+ i for i in docindex]
    
    #initializing a dataframe
    df_clone = pd.DataFrame(zip(docindex2, distindex))

    #renaming columns
    df_clone.columns = ['UID', 'Clone?']
    
    return df_clone
	
	
#CODE TO VISUALIZE DOCUMENT SIMILARITY----------------------------------------------------
#Code to plot a dendrogram of document similarity based on LSA and cosine similarity
#input: the lesson's folder path (for example: 'C:/Users/eabalo/Documents/STAAR2014/4g/405B/')
#output: a dendrogram plot
def lsa_dendrogram(lessonpath):
    #document-term matrix and document indices
    dtm, docindex, lessonname = dtm_matrix(lessonpath)
    
    #reconstructed dtm matrix using LSA and a reduced subspace of dimension 3
    dtm2 = LSA_dtm(dtm, 3)
    
    #distance metric based on cosine similarity
    dist = 1 - cosine_similarity(dtm)

    dist = np.round(dist, 10)

    #linkage matrix
    linkage_matrix = ward(dist)

    #dendrogram
    show(dendrogram(linkage_matrix, orientation='right', labels = docindex))
	
	
	
#Code to plot a scatter plot of document similarity based on LSA, 
# cosine similarity, and multidimensional scaling
#input: the lesson's folder path
#output: returns a scatter plot
def  lsa_mds_plot(lessonpath):
    #document-term matrix and document indices
    dtm, docindex, lessonname = dtm_matrix(lessonpath)
    
    #reconstructed dtm matrix using LSA and a reduced subspace of dimension 3
    dtm2 = LSA_dtm(dtm, 3)
    
    #distance metric based on cosine similarity
    dist = 1 - cosine_similarity(dtm)

    dist = np.round(dist, 10)
    
    #multi-dimensional scaling
    mds = MDS(n_components=2, dissimilarity = 'precomputed', random_state = 1)

    pos = mds.fit_transform(dist)

    xs, ys = pos[:, 0], pos[:, 1]

    for x, y, name in zip(xs, ys, docindex):
        color = 'orange' if 'th' in name else 'skyblue'
        scatter(x, y, c=color)
        plt.text(x, y, name)