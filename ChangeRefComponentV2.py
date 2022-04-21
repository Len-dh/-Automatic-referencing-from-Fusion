#Author-
#Description-

from cProfile import label
from itertools import count
import itertools
from tkinter.tix import Tree

import adsk.core, adsk.fusion, adsk.cam, traceback
import os, sys, datetime
import pandas as pd

#######################################################################################################################

AUTHOR_NAME = 'Lenny DAHO'

#######################################################################################################################



def ImportPandaLibrary(ui):
    install_pandas = sys.path[0] +'\Python\python.exe -m pip install pandas'
    os.system('cmd /c "' + install_pandas + '"')
    
    try:
        import pandas
        return ui.messageBox("Installation succeeded !")
    except:
        return ui.messageBox("Failed when executing 'import pandas'")

#######################################################################################################################


def creerDate(secondes):
	date = datetime.date.fromtimestamp(secondes)
	return str(date)

def getMaterialInfo(bdy :adsk.fusion.BRepBody) -> str:
    return "{}".format(bdy.material.name)


def OccurenceList(occs):
    occur_list = []
    bom = []
    # Get the value of the property.
    
    for occ in occs: 
        #Browse all components of the specific occ
        comp = occ.component 
        occur_list.append(occ.fullPathName)
        secCrea = comp.parentDesign.parentDocument.dataFile.versions.item(len(comp.parentDesign.parentDocument.dataFile.versions)-1).dateCreated
        secLast = comp.parentDesign.parentDocument.dataFile.latestVersion.dateCreated
        dateCrea = creerDate(secCrea)
        dateLast = creerDate(secLast)
        physicalProperties = occ.physicalProperties
        bodies = comp.bRepBodies
        jj = 0
        # for bomI in bom:
        #     if bomI['component'] == comp:
        #         # Increment the instance count of the existing row.
        #         bomI['instances'] += 1
        #         break
        #     jj += 1
        # if jj == len(bom):
        #     # Gather any BOM worthy values from the component
        #     bodies = comp.bRepBodies
        for bodyK in bodies:
            infoMateriau = getMaterialInfo(bodyK)
        #     # get Material Info
        
            
                
        # Add this component to the BOM
        bom.append({
            'component': comp,
            'reference' : comp.partNumber, #Récupération du numéro de la pièce
            'name': comp.name, #Récupération du nom de la pièce
            'description': comp.description, #Récupération de la description de la pièce
            'materiau':infoMateriau, #Récupération du matériau utilisé dans la conception
            'instances': 1,
            'masse' : physicalProperties.mass, #Récupération de la masse de la pièce
            'volume' : physicalProperties.volume, #Récupération du volume de la pièce
            'densite' : physicalProperties.density, #Récupération de la densité du matériau
            'aire' : physicalProperties.area, #Récupération de l'aire de la pièce
            'Date de création' : dateCrea,
            'Version' : comp.parentDesign.parentDocument.dataFile.latestVersionNumber,
            'Date dernière modification' : dateLast,
            'Auteur modification' : AUTHOR_NAME
        })

    return occur_list, bom



##############################################################################################################



def SortedOccurenceList(ui, occur_list, bom):

    FinalOccurenceList = []
    IndexListToBom = []
    SortedBom = []
    nbElemIndexationInPath = 2
    occlist = []

    if not occur_list:
        ui.messageBox("empty list")
    else:
        msg = "the list exist\n"
        for i, occ in enumerate(occur_list):
            if occ.find('+') != -1 and i == 0:
                SplitForAssIndexation = occ.split('+')
                nbElemIndexationInPath = len(SplitForAssIndexation)
                IndexListToBom.append(i)
                FinalOccurenceList.append(occ)
            elif occ.find('+') != -1 and i > 0:
                SplitForAssIndexation = occ.split('+')
                if nbElemIndexationInPath <= len(SplitForAssIndexation):
                    IndexListToBom.append(i)
                    FinalOccurenceList.append(occ)
                    # nbElemToRemove = occ.count('+')
                    nbElemIndexationInPath = len(SplitForAssIndexation)
                elif nbElemIndexationInPath > len(SplitForAssIndexation):
                    nbElemToRemove = occ.count('+') 
                    if nbElemToRemove != 0:
                        nbElemIndexationInPath = len(SplitForAssIndexation)
                        pass
                    else:
                        IndexListToBom.append(i)
                        FinalOccurenceList.append(occ)
                        nbElemIndexationInPath = len(SplitForAssIndexation)
                    nbElemToRemove -= 1
            
            elif occ.find('+') == -1 and i > 1:

                if FinalOccurenceList[-1].count('+') > len(occ):
                    pass
                elif occur_list[i-1].find('+') == -1:  
                    IndexListToBom.append(i)  
                    FinalOccurenceList.append(occ)     
            elif occ.find('+') == -1 and i <= 1:
                IndexListToBom.append(i)
                FinalOccurenceList.append(occ)

        ui.messageBox('occur_list : ' + str(occur_list) + '\n\nFinalOccurenceList : ' + str(FinalOccurenceList))

    for indexToBom in IndexListToBom:
        SortedBom.append(bom[indexToBom])

    ui.messageBox('FinalOccurenceList size : ' + str(len(FinalOccurenceList)) + '\n\noccur_list size : ' + str(len(occur_list)) + '\n\nIndexListToBom : ' + str(len(IndexListToBom)) + '\n\nbom : ' + str(len(bom)) + '\n\nSortedBom : ' + str(len(SortedBom)))


    return FinalOccurenceList, SortedBom



#######################################################################################################################


def OccurToLabel(ui, FinalOccurenceList, SortedBom):
    # Si Element devant un +, Index = ass
    # Si element jamais devant un +, Index = sub_ass
    # Trie de la liste final : Si un element apparait plusieurs fois, et qu'il a 2 index différent, l'index ASS devient prioritaire
    # Dans le trie final, ne garder qu'une seule fois les elements

    msg = msg_2 =''
    SplitForAssIndexation = []
    AssemblyIndex = ''
    SubAssIndex = ''
    n = 0
    AssemblyList = []
    subAssemblyList = []
    FunctionalAreaList = []
    
    PartList = []
    labelList = []
    LabelBom = []

    nbElemOccList = len(FinalOccurenceList)
    if not FinalOccurenceList:
        ui.messageBox("empty list")
    else:
        msg = "the list exist\n"
        for nb_elem in range(nbElemOccList):
           
            if FinalOccurenceList[nb_elem].find('+') != -1:
                SplitForAssIndexation = FinalOccurenceList[nb_elem].split('+')

                for n in range(len(SplitForAssIndexation)):
                    FunctionalAreaIndex = 'N' + str(n)
                    FunctionalAreaList.append(FunctionalAreaIndex)

                    # Prendre la taille de la liste SplitForAssIndexation, tout ce qui se trouve en n-1 de la liste prend en Index Ass 
                    if n < len(SplitForAssIndexation) - 1:
                        msg_2 += str(FinalOccurenceList[nb_elem]) + ' True - There is/are ' + str(FinalOccurenceList[nb_elem].count('+')) + ' + in the list ' + str(SplitForAssIndexation) +'\n'

                        AssemblyIndex = 'ASS' 

                        PartList.append(SplitForAssIndexation[n])
                        AssemblyList.append(AssemblyIndex)
                        subAssemblyList.append('')
                        LabelBom.append(None)
                        #ui.messageBox("n : " + str(n) + " len(SplitForAssIndexation) : " + str(len(SplitForAssIndexation) - 1) + '\n' + str(SplitForAssIndexation[n]) + " = " + LevelAssenblyIndex)
                    else : 
                        msg_2 += str(FinalOccurenceList[nb_elem]) + ' True - There is/are ' + str(FinalOccurenceList[nb_elem].count('+')) + ' + in the list' + str(SplitForAssIndexation) + '\n'
                        SubAssIndex = 'SUB_ASS' 
                        PartList.append(SplitForAssIndexation[n])
                        AssemblyList.append('')
                        subAssemblyList.append(SubAssIndex)
                        LabelBom.append(SortedBom[nb_elem])
                        #ui.messageBox("n : " + str(n) + " len(SplitForAssIndexation) : " + str(len(SplitForAssIndexation) - 1)+ '\n' + str(SplitForAssIndexation[n]) + " = " + LevelAssenblyIndex)    

            else: 
                msg_2 += str(FinalOccurenceList[nb_elem]) + ' False \n'
                SubAssIndex = 'SUB_ASS'
                FunctionalAreaIndex = 'N0'
                #ui.messageBox((FinalOccurenceList[nb_elem]) + ' = ' + LevelAssenblyIndex)
                PartList.append(FinalOccurenceList[nb_elem])
                AssemblyList.append('')
                subAssemblyList.append(SubAssIndex)
                FunctionalAreaList.append(FunctionalAreaIndex)
                LabelBom.append(SortedBom[nb_elem])


    labelList = list(zip(FunctionalAreaList, AssemblyList, subAssemblyList, PartList))

    # ui.messageBox(msg + msg_2)
    ui.messageBox('Label list size : ' +  str(len(labelList)) + '\n\nLabel bom size: ' + str(len(LabelBom))) # + '\n\nLabel bom : ' + str(len(LabelBom))

    #Verifier le sens de labelList par rapport a celui de FinalOccurlist
    
    return labelList, LabelBom



#######################################################################################################################

def removeDuplicatesTuple(lst):
      
    return [t for t in (set(tuple(i) for i in lst))]

def removeDuplicatesTupleWithoutLastValue(ui, lst):
    check_val = set()      #Check Flag
    res = []
    for i in lst:
        if (i[0], i[1], i[2], i[3]) not in check_val:
            res.append(i)
            check_val.add((i[0], i[1], i[2], i[3]))

    return res

def GetListToFinalList1(ui, labelList, LabelBom):
    SortedLabelList = []
    functionalAreaList = []
    assemblyList = []
    subAssemblyList= []
    partList = []
    GetListToFinalList = []
    GetListToFinalListPart1 = []
    i_list = []
    TreeBom1 = []


    for i, (functionalArea, assembly, subAssembly, part) in enumerate(labelList):
        if assembly == 'ASS' and labelList[i-1][1] != 'ASS':
            functionalAreaList.append(functionalArea)
            assemblyList.append(assembly)
            subAssemblyList.append(subAssembly)
            partList.append(part)
            i_list.append(i)
            TreeBom1.append(LabelBom[i])
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 0: ' + str(SortedLabelList))            
        elif assembly == 'ASS' and labelList[i-1][1] == 'ASS':
            if not SortedLabelList:
                continue 
            else: 
                GetListToFinalList.append(SortedLabelList)
            SortedLabelList = []
            functionalAreaList = []
            assemblyList = []
            subAssemblyList= []
            partList = [] 
            i_list = []
            TreeBom1 = []
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 1: ' + str(SortedLabelList) + '\n\nGetListToFinalList: ' + str(GetListToFinalList))            
            continue
        elif subAssembly == 'SUB_ASS' and labelList[i-2][1] == 'ASS' and int(labelList[i-2][0][1]) == int(functionalArea[1]) -2:
            SortedLabelList = []
            functionalAreaList = []
            assemblyList = []
            subAssemblyList= []
            partList = []     
            i_list = []
            TreeBom1 = []
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 2: ' + str(SortedLabelList))            
            continue
        elif subAssembly == 'SUB_ASS' and labelList[i-1][1] == 'ASS' and int(labelList[i-1][0][1]) == int(functionalArea[1]) -1:
            functionalAreaList.append(functionalArea)
            assemblyList.append(assembly)
            subAssemblyList.append(subAssembly)
            partList.append(part)    
            i_list.append(i)    
            TreeBom1.append(LabelBom[i])
            SortedLabelList = list(zip(functionalAreaList, assemblyList, subAssemblyList, partList, i_list))
            if i == len(labelList) - 1:
                GetListToFinalList.append(SortedLabelList)
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 3: ' + str(SortedLabelList) + '\n\nGetListToFinalList: ' + str(GetListToFinalList))
        elif assembly == 'ASS' and labelList[i-2][1] == 'ASS' and labelList[i-2][3] == part and int(labelList[i-1][0][1]) == int(functionalArea[1])+1:
            functionalAreaList.append(functionalArea)
            assemblyList.append(assembly)
            subAssemblyList.append(subAssembly)
            partList.append(part)   
            i_list.append(i)  
            TreeBom1.append(LabelBom[i])  
            SortedLabelList = list(zip(functionalAreaList, assemblyList, subAssemblyList, partList, i_list))
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 4: ' + str(SortedLabelList))

        elif subAssembly == 'SUB_ASS' and labelList[i-3][1] == 'ASS' and int(labelList[i-3][0][1]) == int(functionalArea[1]) -1:
            functionalAreaList.append(functionalArea)
            assemblyList.append(assembly)
            subAssemblyList.append(subAssembly)
            partList.append(part)  
            i_list.append(i)  
            TreeBom1.append(LabelBom[i])  
            SortedLabelList = list(zip(functionalAreaList, assemblyList, subAssemblyList, partList, i_list))
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 5: ' + str(SortedLabelList))

        else:
            if not SortedLabelList:
                continue 
            else: 
                GetListToFinalList.append(SortedLabelList)
            SortedLabelList = []
            functionalAreaList = []
            assemblyList = []
            subAssemblyList= []
            partList = []      
            i_list = []
            TreeBom1 = []
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 6: ' + str(SortedLabelList) + '\n\nGetListToFinalList: ' + str(GetListToFinalList))

    for glist in GetListToFinalList:
        glist = removeDuplicatesTupleWithoutLastValue(ui, glist)
        sortedglist = sorted(glist, key=lambda tup: tup[1], reverse=True)
        GetListToFinalListPart1.extend(sortedglist)
    return GetListToFinalListPart1, TreeBom1

def GetListToFinalList2(ui, labelList, LabelBom):
    SortedLabelList = []
    functionalAreaList = []
    assemblyList = []
    subAssemblyList= []
    partList = []
    GetListToFinalList = []
    GetListToFinalListPart2 = []
    i_list = []
    TreeBom2 = []
    
    for i, (functionalArea, assembly, subAssembly, part) in enumerate(labelList):
        if assembly == 'ASS' and labelList[i-1][1] == 'ASS' : # and labelList[i-1][3] != part and labelList[i-1][0] == functionalArea -1
            functionalAreaList.append(functionalArea)
            assemblyList.append(assembly)
            subAssemblyList.append(subAssembly)
            partList.append(part)
            i_list.append(i)
            TreeBom2.append(LabelBom[i])
            SortedLabelList = list(zip(functionalAreaList, assemblyList, subAssemblyList, partList, i_list))
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 7: ' + str(SortedLabelList))            
        elif subAssembly == 'SUB_ASS' and labelList[i-1][1] == labelList[i-2][1] ==  'ASS' and labelList[i-1][3] == labelList[i-4][3] and int(labelList[i-1][0][1]) == int(labelList[i-4][0][1]) == int(functionalArea[1]) -1: # and int(labelList[i-3][0][1]) == int(functionalArea[1]) -1
            functionalAreaList.append(functionalArea)
            assemblyList.append(assembly)
            subAssemblyList.append(subAssembly)
            partList.append(part)   
            i_list.append(i) 
            TreeBom2.append(LabelBom[i])
            SortedLabelList = list(zip(functionalAreaList, assemblyList, subAssemblyList, partList, i_list))
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 8: ' + str(SortedLabelList)) 
        elif subAssembly == 'SUB_ASS' and labelList[i-1][1] == 'ASS' and labelList[i-2][1] == 'ASS' and int(labelList[i-1][0][1]) == int(functionalArea[1]) -1:
            functionalAreaList.append(functionalArea)
            assemblyList.append(assembly)
            subAssemblyList.append(subAssembly)
            partList.append(part)    
            i_list.append(i)
            TreeBom2.append(LabelBom[i])
            SortedLabelList = list(zip(functionalAreaList, assemblyList, subAssemblyList, partList, i_list))
            # ui.messageBox('label list : ' + str(labelList) + '\n\nSortedLabelList 9: ' + str(SortedLabelList)) 
        elif assembly == "ASS" and labelList[i-3][1] == 'ASS' :   
            continue
        else:
            if not SortedLabelList:
                continue 
            else: 
                GetListToFinalList.append(SortedLabelList)        
            SortedLabelList = []
            functionalAreaList = []
            assemblyList = []
            subAssemblyList= []
            partList = []   
            i_list = [] 
            TreeBom2 = []        

    for glist in GetListToFinalList:
        glist = removeDuplicatesTupleWithoutLastValue(ui, glist)
        sortedglist = sorted(glist, key=lambda tup: tup[1], reverse=True)
        GetListToFinalListPart2.extend(sortedglist)
    return GetListToFinalListPart2, TreeBom2

def GetListToFinalList3(ui, labelList, LabelBom):
    GetListToFinalListPart3 = []
    functionalAreaList = []
    assemblyList = []
    subAssemblyList= []
    partList = []  
    GetListToFinalListPart3 = []
    i_list = []
    TreeBom3 = []

    for i, (functionalArea, assembly, subAssembly, part) in enumerate(labelList):
        if subAssembly == 'SUB_ASS' and functionalArea == 'N0':
            functionalAreaList.append(functionalArea)
            assemblyList.append(assembly)
            subAssemblyList.append(subAssembly)
            partList.append(part) 
            i_list.append(i)
            TreeBom3.append(LabelBom[i])
        else: 
            continue

        GetListToFinalListPart3 = list(zip(functionalAreaList, assemblyList, subAssemblyList, partList, i_list))
    return GetListToFinalListPart3, TreeBom3

def ListFusion(ui, GetListToFinalListPart1, GetListToFinalListPart2, GetListToFinalListPart3, TreeBom1, TreeBom2, TreeBom3):
    TreeList = GetListToFinalListPart1 + GetListToFinalListPart2 + GetListToFinalListPart3
    TreeList = sorted(TreeList, key=lambda tup: tup[4])
    TreeBom = TreeBom1 + TreeBom2 + TreeBom3
    
    return TreeList


def AddTreeList(ui, FinalOccurenceList, labelList, LabelBom): #add IndexList

    # si Part apparait plusieur fois, Ne garder l'element qu'une seule fois en gardant le 2ème argument NX avec le X le plus eleve pour le niveau de zone
    # si Part apparait plusieur fois, Ne garder l'element qu'une seule fois en gardant le 2ème argument en tant que ASS
    # faire un test pour savoir si len(newZippedList) == len(occur_list) 

    TreeBom = []

    ui.messageBox('The zippedList is : ' + str(labelList))

    GetListToFinalListPart1, TreeBom1 = GetListToFinalList1(ui, labelList, LabelBom)
    ui.messageBox('GetListToFinalListPart1 : ' + str(GetListToFinalListPart1))

    GetListToFinalListPart2, TreeBom2 = GetListToFinalList2(ui, labelList, LabelBom)
    ui.messageBox('GetListToFinalListPart2 : ' + str(GetListToFinalListPart2))

    GetListToFinalListPart3, TreeBom3 = GetListToFinalList3(ui, labelList, LabelBom)
    ui.messageBox('GetListToFinalListPart3 : ' + str(GetListToFinalListPart3))

    TreeList = ListFusion(ui, GetListToFinalListPart1, GetListToFinalListPart2, GetListToFinalListPart3, TreeBom1, TreeBom2, TreeBom3)
    ui.messageBox('TreeList size: ' + str(len(TreeList)) + ' - occur_list size: ' + str(len(FinalOccurenceList)))
    ui.messageBox('Label list : ' + str(labelList) + '\n\nTreeList : ' + str(TreeList))


    assert TreeList != FinalOccurenceList, ui.messageBox('Error, Tree list doen\'t match !')
    assert TreeBom != FinalOccurenceList, ui.messageBox('Error, Tree list doen\'t match !')

    return TreeList



#######################################################################################################################



def LabelToNumeriacalIndex(ui, TreeList):
    faiVar = aiVar = saiVar = partVar = faiVarPred = 0
    TreeListToIndex = []
    lastpart = ''

    for i, (functionalareaindex, assemblyindex, subassemblyindex, part, index) in enumerate(TreeList):
        # Firdt case : Part 

        faiVar = 0

        if assemblyindex == '' and subassemblyindex == 'SUB_ASS' and i == 0 :
            aiVar = 0
            saiVar = 0
            partVar += 1
            TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
   
            TreeListToIndex.append(TreeList[i])  


        elif assemblyindex == 'ASS' and subassemblyindex == '' and i == 0:
            aiVar += 1
            saiVar = 0
            partVar = 0
            TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")

            TreeListToIndex.append(TreeList[i])
        
        # != 0 
        if assemblyindex == '' and subassemblyindex == 'SUB_ASS'  and i != 0 :

        
            if int(TreeListToIndex[len(TreeListToIndex)-1][1]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 :
                aiVar = 0
                saiVar = 0
                partVar += 1
                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")    
                TreeListToIndex.append(TreeList[i]) 
                lastpart = partVar

            elif int(TreeListToIndex[len(TreeListToIndex)-1][0]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) == 0:
                aiVar = int(TreeListToIndex[len(TreeListToIndex)-1][1])
                saiVar += 0
                partVar = 1  
                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[i])

            elif int(TreeListToIndex[len(TreeListToIndex)-1][0]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 :
                aiVar = int(TreeListToIndex[len(TreeListToIndex)-1][1])
                saiVar += 0
                partVar += 1  
                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[i])

            elif int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) == 0:
                aiVar = int(TreeListToIndex[len(TreeListToIndex)-1][1])
                saiVar = int(TreeListToIndex[len(TreeListToIndex)-1][2])
                partVar = 1 
                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[i])

            elif int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 :
                aiVar = int(TreeListToIndex[len(TreeListToIndex)-1][1])
                saiVar = int(TreeListToIndex[len(TreeListToIndex)-1][2])
                partVar += 1 
                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[i])

            elif int(TreeListToIndex[len(TreeListToIndex)-1][0]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) == 0:
                faiVar = int(TreeListToIndex[len(TreeListToIndex)-1][0])
                aiVar = int(TreeListToIndex[len(TreeListToIndex)-1][1])
                saiVar = 0
                partVar += 1 
                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[i]) 

            elif int(TreeListToIndex[len(TreeListToIndex)-1][0]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 and functionalareaindex == 'N1':
                faiVar = int(TreeListToIndex[len(TreeListToIndex)-1][0])
                aiVar = int(TreeListToIndex[len(TreeListToIndex)-1][1])
                saiVar = int(TreeListToIndex[len(TreeListToIndex)-1][2])
                partVar += 1 
                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[i])

            elif int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 and functionalareaindex == 'N0':
                aiVar = 0
                saiVar = 0
                partVar = int(lastpart) + 1 

                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")    
                TreeListToIndex.append(TreeList[i]) 
                lastpart = partVar


            elif int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 and functionalareaindex == 'N0':
                aiVar = 0
                saiVar = 0
                partVar = int(lastpart) + 1 

                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")    
                TreeListToIndex.append(TreeList[i]) 
                lastpart = partVar

            else:
                'Error !'

        elif assemblyindex == 'ASS' and subassemblyindex == '' and i != 0:

            if int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 and faiVarPred == 0:
                aiVar = 1
                saiVar = 0
                partVar = 0      
                faiVar = 1     
                faiVarPred = faiVar

                TreeList[len(TreeListToIndex)] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[len(TreeListToIndex)])

            if int(TreeListToIndex[len(TreeListToIndex)-1][1]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 and faiVarPred > 0:
                aiVar = 1
                saiVar = 0
                partVar = 0      
                faiVar = faiVarPred + 1    
                faiVarPred = faiVar

                TreeList[len(TreeListToIndex)] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[len(TreeListToIndex)])

            if int(TreeListToIndex[len(TreeListToIndex)-1][1]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 :
                aiVar = 1
                saiVar = 0
                partVar = 0  
                TreeList[i] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[i])

            if int(TreeListToIndex[len(TreeListToIndex)-1][1]) > 0 and int(TreeListToIndex[len(TreeListToIndex)-1][2]) == 0 and int(TreeListToIndex[len(TreeListToIndex)-1][3]) > 0 :
                aiVar = int(TreeListToIndex[len(TreeListToIndex)-1][1])
                saiVar = 1
                partVar = 0  
                TreeList[len(TreeListToIndex)] = (f"{faiVar:04}", f"{aiVar:04}", f"{saiVar:04}", f"{partVar:04}")
                TreeListToIndex.append(TreeList[len(TreeListToIndex)])




    ui.messageBox(str(TreeList))
        

    return TreeList

# Impossible to use Pandas in Fusion 360. So impossible to create CSV file in python with fusion 360 app

def CreateReferenceDataframe(ui):
    # use replace to replache Label by int index element 
    try:
        referenceDataframe = pd.DataFrame(columns = ['ProjectName', 'ProjectVersion', 'FunctionalAreaIndicator', 'AssemblyIndicator', 'SubAssemblyIndicator', 'PartNumber'])
        ui.messageBox("Dataframe is created !")
        return referenceDataframe
    except:
        return ui.messageBox("Failed to create dataframe !")




def AddDataInDataframe(ui, referenceDataframe, numeriacalIndexList):
    ui.messageBox(str(numeriacalIndexList))

    if not numeriacalIndexList:
        ui.messageBox('the numeriacalIndexList  doesn\'t exist')
    try: 
        for num_index in numeriacalIndexList:
            referenceDataframe = referenceDataframe.append({'ProjectName' :  'AssemblageTestN2',
                                                            'ProjectVersion' : 'v18',
                                                            'FunctionalAreaIndicator' : num_index[0],
                                                            'AssemblyIndicator' : num_index[1],
                                                            'SubAssemblyIndicator' : num_index[2],
                                                            'PartNumber' : num_index[3]}, ignore_index=True)

        ui.messageBox("Reference dataframe complete !\n" + str(referenceDataframe))

        return referenceDataframe

    except:
        return ui.messageBox("Failed to create dataframe !")


#######################################################################################################################


def DataframeToCSV(ui, referenceDataframe):
    #Convert Dataframe to CSV 
    try:
        if os.path.exists('C:\\Users\\ldaho\\Referencev3.csv'):
            ui.messageBox('The csv file exist')
        else:
            referenceDataframe.to_csv('C:\\Users\\ldaho\\Referencev3.csv')
            return ui.messageBox('The csv is created. the path is : ' + os.path.abspath("Referencev3.csv"))
    except:
        ui.messageBox('Failed to create Referencev3.csv') # C:\Users\ldaho\AppData\Local\Autodesk\webdeploy\production\19107935ce2ad08720646cb4a31efe37d8a5f41b\ReferenceCsv.csv
    # Faire en sorte que si le CSV existe, ne par supprimer ou remplacer les données. Juste ajouter les nouevlles donnéees



#######################################################################################################################


def AddCsvfileInOriginalCSV(ui):
    return 0 

def CsvToFusion(ui):
    return 0


def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        design = app.activeProduct
        root = design.rootComponent
        occs = root.allOccurrences

        # ImportPandaLibrary(ui)

        occur_list, bom = OccurenceList(occs)

        ui.messageBox('bom : ' + str(bom[7]))

        FinalOccurenceList, SortedBom = SortedOccurenceList(ui, occur_list, bom)

        labelList, LabelBom = OccurToLabel(ui, FinalOccurenceList, SortedBom)

        TreeList = AddTreeList(ui, FinalOccurenceList ,labelList, LabelBom)

        numeriacalIndexList = LabelToNumeriacalIndex(ui, TreeList)

        # referenceDataframe = CreateReferenceDataframe(ui)

        # referenceDataframe = AddDataInDataframe(ui, referenceDataframe, numeriacalIndexList)

        # DataframeToCSV(ui, referenceDataframe)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
