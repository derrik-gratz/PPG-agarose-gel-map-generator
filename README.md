# agarose-gel-map-generator

Sample input: 'gel_02-19-2020-06-03-38.xlsx'. 
Sample output: '2020-02-19 Sec_2%_E-GEL_Map.xlsx', '2020-02-19 Sec_4%_E-GEL_Map.xlsx', and'2020-02-19 Sec_Agarose_Gel_Map.xlsx'

Takes lists of what samples need which tests. Organizes the samples by test in corresponding groups with controls.
These groups are output to the gel map templates, which mirror how the samples will be loaded into agarose gels for genotyping.
This allows us to keep track of samples/tests and where they are in the gels. 
Later, genotypes can be entered in the gel maps as techs at PPG are reviewing gel images and making genotype calls. 
Since PPG uses two-step genotype verification, some tests are run twice in our group, so separate results are compared with the
discordance program to make sure genotype calls converge.

We run 3 types of gels so 3 gel maps are made. 

the 'Sec_2%_E-GEL_Map.xlsx', 'Sec_4%_E-GEL_Map.xlsx', and 'Sec_Agarose_Gel_Map.xlsx' serve as blank templates for the program to add
samples and tests to. 

The map making key keeps track of unique loading patterns for certain tests that will determine how samples are arranged on the map.


