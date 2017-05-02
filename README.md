# Synopsis
**MACRO_Network (Santosh)_Jun25.1.xlsm**
Excel macro to compute basic social network indices (e.g. centralization, centrality) for individuals and groups.

# Motivation
This is part of the tools/code snippets that were developed during my RA-ship at ISB under Ruchi Sinha, PhD to automate some of the steps related to micro/macro organizational research methods.

# Contributor
Santosh Srinivas, PhD Student, University of Texas at Austin

# Usage Guidelines
## Prepare the Input Worksheets in the Excel as follows:
- SheetName: DATA
  * This sheet has the data for which the SNA has to be computed
  * First row: HEADER
  * First column:  Unique identifier for member who is rating
  * Second column: Unique identifier for the group that the rater belongs to.
  * Third to N column: Members rated by the rater. Number of columns should be maximum group size -1
  * N+1 onwards: Survey item rating score. Rating in N+1 column will be for the member specified in 3rd column, in N+2 column for the member in 4th column, and so on. Number of columns should be maximum group size  - 1.
  * Note that the data need not be sorted by group or individual IDs. The macro takes care of this.
  * Note that the order of rated member IDs need not be same across raters in the same group. The macro takes care of this.
  * Note that name of the sheet has to be "DATA"
  * This sheet is password protected (password: "macro") to avoid inadvertent edits
- SheetName: StudentList
  * This sheet has complete list of people on whom survey was administered. 
  * First row: HEADER
  * First column:  Unique identifier for member who is rating
  * Second column: Unique group identifier for the rater
  * Note that the data need not be sorted by group or individual IDs. The macro takes care of this.
  * Note that name of the sheet has to be "StudentList"
  * This sheet is password protected (password: "macro") to avoid inadvertent edits
- SheetName: Instructions
  * The current sheet that has compute button , which when clicked opens a form for SNA computation
  * Note that name of the sheet has to be "INSTRUCTIONS"
  * This sheet is password protected (password: "macro") to avoid inadvertent edits
## Compute the social network indices
- Once the input data is specified in the format detailed above, click the button 'Compute Network Indices' in the "Instructions" worksheet in the Excel macro for computing network indices
  * Keep a note of status bar when the macro runs. It provides progress of network index computation. 
  * The macro takes anywhere from 1 to 15 min for about 800-1000 rows of DATA. On an Intel i7 processor the bechmark speed for this data size, with all network computation methods (7 methods) choosen at once is about 1.5- 2 min.
  * All sheets are protected with PASSWORD "macro" to avoid any inadvertent typing overs. Please free free to unlock them as needed (but please note that changing the password from "macro" to anything else will render the software unusable).
## Check the Results/Output Worksheets in the Excel
- SheetName: Raw
  * This sheet has SNA computations based on RAW matrix (i.e. no imputations, reconstructions or complete case). 
  * The whole matrix is considered as is, with all non-responses left as is.
  * Note that for SNA computations (e.g. Density), original group size (irrespective of how many responded in that group) is taken as n. This n is <= max group size.
- SheetName: CC
  * This sheet has SNA computations based on complete-case approach.
  * The complete-case approach, known also as ‘listwise’ deletion of actors (Huisman and Steglich, 2008), removes not only the rows for the non-respondents but also their columns.
  * Note that for SNA computations (e.g. Density),  group size is number of members in the complete-case matrix for the group. This is <= actual group size.
- SheetName: RC
  * This sheet has SNA computations based on reconstruction approach.
  * Reconstruction of the missing outgoing ties of non-respondents occurs when they are replaced by the observed incoming ties to those actors. 
  * Note that  for two non-respondents the reconstruction of ties between them is not possible. It is left as empty.
  * Note that for SNA computations, original group size (irrespective of how many responded in that group) is taken as n
- SheetName: RCIM
  * This sheet has SNA computations based on reconstruction + modal imputation approach.
  * This approach combines the reconstruction procedure with imputations based on modal indegree values for ties between nonrespondents.
  * Note that when mode of incoming ties is both 0 and 1, then based on whether conservative/non-conservative approach is chosen, 0/1 is imputed respectively.
  * Note that for SNA computations, original group size (irrespective of how many responded in that group) is taken as n
- SheetName: IMNT
  * This sheet has SNA computations based on null-tie imputations.
  * In this approach zeroes are imputed instead of missing ties (irrespective of whether the network is dense or not)
  * Note that for SNA computations, original group size (irrespective of how many responded in that group) is taken as n
- SheetName: IMD
  * This sheet has SNA computations based on total mean of observed ties.
  * For binary network this total mean is same as density of the network.
  * Threshold value is 0.5; This implies that if density of network is > 0.49 then 1 is imputed, else 0 is imputed.
  * Note that for SNA computations, original group size (irrespective of how many responded in that group) is taken as n
- SheetName: IMIM
  * This sheet has SNA computations based on item mean (i.e. mean of incoming ties)
  * For binary networks this implies imputing 1 if actors are popular given their received ties
  * Threshold value is 0.5; This implies that if mean of incoming tie for j is > 0.49 then 1 is imputed for missing tie Xij, else 0 is imputed.
  * Note that for SNA computations, original group size (irrespective of how many responded in that group) is taken as n
- SheetName: SNA
  * This sheet has the final computations (with header row) that could be imported to SPSS.
  * Basically this has merged output sheets (RAW, CC, RC, RCIM, IMNT, IMD and IMIM)
- SheetName: RawMatrix
  * This is the matrix output of DATA sheet. 
- SheetName: --DichotMatrix
  * These are dichotomizations of matrix meant for network index calculations
- SheetName: T--DichotMatrix
  * These are transposed matrices for those in the corresponding dichotomization sheets
# Final Notes
This software has been tested to the best of my ability. However, I assume no liability, direct or otherwise, which may result from the use of this software. The software is provided for free and as-is. If you use this software, it would be great to hear from you at santosh.b.srinivas@gmail.com. Any acknowledgement by means of a citation where appropriate would be highly appreciated as well. 





