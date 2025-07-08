# run_all.py

import A_Page_Separator as A
import B_Header_Finder as B
import C_Header_Separator as C
import D_Actual_and_Goal_Merger as D

if __name__ == "__main__":
    print("▶️ Running A_Page_Separator")
    A.main()
    
    print("▶️ Running B_Header_Finder")
    B.main()
    
    print("▶️ Running C_Header_Separator")
    C.main()
    
    print("▶️ Running D_Actual_and_Goal_Merger")
    D.main()
