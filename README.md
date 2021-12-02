# Optimization-milk-shelf-life
This folder contains two optimization codes for FFAR pasteurized milk project.

Paper: Optimizing pasteurized fluid milk shelf-life through microbial spoilage reduction

Authors: Forough Enayaty-Ahangar (corresponding author), Sarah Murphy, Nicole Martin, Martin Wiedmann, and Renata Ivanek

This paper was published in the Frontiers journal (Front. Sustain. Food Syst., 08 July 2021 | https://doi.org/10.3389/fsufs.2021.670029 ).

There are two Python codes in this folder for the two optimization models:  

1) "MPBOP.py" file:  this code has the model for the milk processor budget optimization problem (MBBOP)
2) "MPSLOP.py" file: this code has the model for the milk shelf-life optimization problem (MSLOP)

They read from 24 input data files (8 small, 8 medium, and 8 large processor instances) and then solve the problems using the Gurobi solver. 

If you have any difficulties running the codes, please contact the corresponding author at forough.enayaty@gmail.com.
