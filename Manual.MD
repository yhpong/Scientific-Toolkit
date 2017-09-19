---
### cPCA.cls
---
This class is used to run principal component analysis.

Requires: [modMath.bas](Modules/modMath.bas)

#### Available methods:

```Ruby
PCA(x() As Double, Optional first_n As Long = 0, Optional use_SVD As Boolean = False, Optional save_rawdata As Boolean = False)
```

**Desc:** Perform PCA transformation on data x()

**Input:** 
- x(),          a NxD array of data, where N is number of samples and D is dimension of data  
- first_n,      number of PCs to calculate. If left at zero, all componets are calculated.  
                      Otherwise only first_n PCs are calculated using power iteration.  
- use_SVD,      if set to TRUE, singular value decompositon is performed on x(). Otherise  
                      covariance matrix of x() is calculated and Jacobi method is used.  
- save_rawdata, if set to TRUE then x() will be saved in class object and can be accessed by method .raw_data  

```Ruby
Reset()
```

**Desc:**   Release memory in class object

```Ruby
Vec(Optional n_vec as variant) As Double()
```

**Desc:**   Read only property, return the first n_vec unit-vectors of PCs, in an array of size Vec(1:D, 1:n_vec)

**Input:**  n_vec, number of vectors to return. If left blank then all calculated eigenvectors are returned

```Ruby
Val(Optional n_vec as variant) As Double()
```

**Desc:**   Read only property, return the first n_vec eigenvalues, in an array of size Val(1:n_vec)

**Input:**  n_vec, number of values to return. If left blank then all calculated eigenvalues are returned

```Ruby
x_PCA(Optional n_dim As Variant) As Double()
```

**Desc:**  Read only property, return projections of raw data on the first n_dim PCs, in an array of size(1:N, 1:n_dim)

**Input:** n_dim, number of components to return. If left blank then all calculated PCs are returned

```Ruby
BiPlot_Print(vRng As Range, Optional PC1 As Long = 1, Optional PC2 As Long = 2, Optional magnify As Double = 1)
```

**Desc:**   output data necessary to chart a biplot in Excel. First two column are scatter chart,
        next two columns are chart with scatter plot with straight lines.
        
**Input:**
 - vRng,       upper-left cell of output range
 - PC1 & PC2,  components to print on the x & y axis respectively, default are the 1st and 2nd components.
 - magnify,    default value is 1 and a vector is shown at length equals to square root of its eigenvalue.
                    The length can be mutiplied by magnify so it can be chart together with x_PCA().




---
### ctSNE.cls
---
This class is used to run t-SNE (t-Distributed Stochastic Neighbor Embedding). Main reference includes:
- "Visualizing High-Dimensional Data Using t-SNE", Laurens van der Maaten (2008)
- "Accelerating t-SNE using Tree-Based Algorithms", Laurens van der Maaten (2014)
- Implementations can be found on the author's [page](https://lvdmaaten.github.io/tsne/)

Requires: [mkdTree](Modules/mkdtree.bas), [cqtree](Modules/cqtree.cls), [cqtree_point](Modules/cqtree_point.cls)

#### Available methods:

```Ruby
tSNE(x() As Double, tgt_dimension As Long, _
            Optional perplexity As Double = 30, Optional perp_err As Double = 0.0001, _
            Optional learn_rate As Double = 100, Optional momentum As Double = 0.5, Optional max_iterate As Long = 1000, _
            Optional input_dist As Boolean = False)
```
**Desc:** Project raw data x() on tgt_dimension.

**Input:** 
- x(),          a NxD array of data, where N is number of samples and D is dimension of data  
- tgt_dimension,  number of output deimsnion, usually it's 2.
- perplexity
- perp_err
- learn_rate
- momentum
- max_iterate
- input_dist, set to TRUE if x() is already a pairwsie distance matrix of size NxN

```Ruby
tSNE_BarnesHut(x() As Double, tgt_dimension As Long, _
            Optional perplexity As Double = 30, Optional perp_err As Double = 0.0001, _
            Optional learn_rate As Double = 100, Optional momentum As Double = 0.5, Optional max_iterate As Long = 1000, _
            Optional input_dist As Boolean = False)
```
**Desc:** Project raw data x() on tgt_dimension. Use tree to accelerate the process. Suitable for large number of samples.

**Input:**  Same as tSNE.

```Ruby
Reset()
```
**Desc:** Release memory from class object

```Ruby
raw_data() As Double()
```
**Desc:** Read only property, return raw data that was saved in class

```Ruby
cost_function(Optional show_index As Boolean = False) As Double()
```
**Desc:** Read only property, return cost fucntion when tSNE was run

**Input:** show_index, if set to false, only return a vector. If set to false, return a 2D array with iterate number in first column.

```Ruby
Output() As Double()
```
**Desc:** Read only property, return porjections as an array of size (1:N, 1:tgt_dimension)


---
### ckMeanCluster.cls
---
This class is used to run k-Means clustering. [k++ scheme](https://en.wikipedia.org/wiki/K-means%2B%2B) is implemented for initialization.

#### Available methods:

```Ruby
kMean_Clustering(x() As Double, k As Long, Optional iterate_max As Long = 100, Optional strType As String = "EUCLIDEAN")
```
**Desc:** Perform k-Means clustering on raw data x().

**Input:** 
- x(),            a NxD array of data, where N is number of samples and D is dimension of data  
- k, target number of clusters
- iterate_max, maximum allowed number of iterations
- strType, type of distance metrics to use. Supports "EUCLIDEAN", "CORREL"

```Ruby
raw_data() As Double()
```
**Desc:** Read only property, returns originald data.

```Ruby
k_cluster() As Long
```
**Desc:** Read only property, returns number of clusters.

```Ruby
cluster_mean() As Double()
```
**Desc:** Read only property, returns mean of each cluster as an array of size (1:k,1:D)

```Ruby
cluster_size() As Long()
```
**Desc:** Read only property, returns size of each cluster as an integer vector of size (1:k)

```Ruby
x_cluster() As Long()
```
**Desc:** Read only property, returns clsuter index of each data point as an integer vector of size (1:N)

```Ruby
Reset()
```
**Desc:** Release memory from class



---
### cGraphAlgo.cls
---
This class is used to build undirected graph, using edge list as the main data structure.
For a graph with m edges, 
- pEdgeList(): a m x 2 array storing the start and end node index of each edge
- pEdgeDist(): a m length vector storing the length of each edge

#### Available methods:

```Ruby
MST_Build(x_dist() As Double)
```
**Desc:** Build Minimum Spanning Tree (MST) from a pairwise distance matrix x_dist().

**Input:** x_dist(),  NxN pairwise distance matrix for N nodes.

```Ruby
PMFG_Build(x_dist() As Double)
```
**Desc:** Build Planar Maximally Filterd Graph (PMFG) from a pairwise distance matrix x_dist().

**Input:** x_dist(),  NxN pairwise distance matrix for N nodes.


```Ruby
Init(EdgeList() As Long, EdgeDist() As Double, node_pos() As Double)
```
**Desc:** Build graph from a previously saved edge list, edge distance and node layout

**Input:** 
- EdgeList(), integer matrix of size m x 2, storing the ending points of m edges.
- EdgeDist(), real vector of size m, storing the length of each edge.
- node_pos(), real matrix of size n x 2, storing the 2D layout of n nodes.

```Ruby
Copy(g As cGraphAlgo)
```
**Desc:** Copy all attribules from anotehr graph 

**Input:** g, a cGraphAlgo object

```Ruby
Reset()
```
**Desc:** Release memory from class

Read-Only Properties:
```
Size() As Long                'number of nodes
n_edge() As Long              'number of edges
EdgeList() As Long()          'edge list, array of size (1:n_edge, 1:2)
EdgeDist() As Double()        'edge length, vector of length (1:n_edge)
node_degree() As Long()       'degree of each node, vector of length (1:Size)
node_degree_wgt() As Double() 'weighted degree of each node, vector of length (1:Size)
node_closeness() As Double()  'closeness of each node, vector of length (1:Size)
node_eigen() As Double()      'eigenvector centrality of each node, vector of length (1:Size)
```

Read-Write properties
```  
node_pos(x() As Double)     'return or assign 2-D layout to nodes, array of size (1:Size, 1:2)
```

```Ruby
Print_Edges() as Variant
```
**Desc:** Return an array of size 3 x n_edge which is used to show edges in Excel on scatter chart.

```Ruby
ForceDirectedLayout(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 500)
```
**Desc:** Force directed algorithm to optimize graph layout, which is accessed by 'node_pos'.  Using quadratic repulsive force between nodes and log-linear spring.

**Input:**
- c1, spring constrant
- c2, spring length
- c3, repulsive force constant
- iter_max, maximum number of iterations allowed.

```Ruby
ForceDirectedLayout_BarnesHut(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 500)
```
**Desc:** Force directed algorithm accelerated with Barnes-Hut algroithm using quad tree.

**Input:** Same as `ForceDirectedLayout`

```Ruby
ForceDirected_MultiLevel(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 300)
```
**Desc:** Force directed algorithm accelerated with Barnes-Hut algroithm and multilevel

**Input:** Same as `ForceDirectedLayout`

```Ruby
Find_degree()
```
**Desc:** Calculate degree of each node which can then be accessed by 'node_degree'

```Ruby
Find_degree_wgt()
```
**Desc:** Calculate weighted degree of each node which can then be accessed by 'node_degree_wgt'. Weight is 1/distance to adjacent node.

```Ruby
Find_closeness()
```
**Desc:** Calculate closenss of each node which can then be accessed by 'node_closeness'. Using Dijkstra's Algorithm.

```Ruby
Find_Eigen(Optional iter_max As Long = 10000, Optional tolerance As Double = 0.0000000001)
```
**Desc:** Calculate eigen-centrality of each node which can then be accessed by 'node_eigen'. Using power iteration.

**Input:** iter_max and tolerance are the maximium iteration of tolerance in using power iteration.
