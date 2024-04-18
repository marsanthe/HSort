
using namespace System.Collections.Generic

$ErrorActionPreference = "Stop"

# .txt in case Manga folder contains .txt with additional info.
$ImageFiles_Array = ('.jpg','.jpeg','.png','.txt')
$global:IsImage_Set = [HashSet[string]]::new()

for($i = 0; $i -le ($ImageFiles_Array.length -1); $i += 1){
    $null = $IsImage_Set.Add($ImageFiles_Array[$i])
}

$IsArchive_Array = ('.zip','.rar','.cbz','.cbr')
$global:IsArchive_Set = [HashSet[string]]::new()

for($i = 0; $i -le ($IsArchive_Array.length -1); $i += 1){
    $null = $IsArchive_Set.Add($IsArchive_Array[$i])
}

$global:Graph = @{}

function New-Graph{

    Param(
        [Parameter(Mandatory)]
        [string]$SourceDir
    )

    Write-Output "Scanning folders. Please wait..."

    $Graph.Clear()

    # Get all nodes (folders).
    # This does NOT include the root directory iteself.
    $AllNodes_ObjectArray = (Get-ChildItem -LiteralPath $SourceDir -Directory -Recurse)

    ### Begin: Insert Root Node ###

    # Initialize new
    $Children_Lst = [List[object]]::new()
    $LeafNode_Lst = [List[object]]::new()
    
    # Get all folders located at root.
    $Children_ObjArr = (Get-ChildItem -LiteralPath $SourceDir -Directory)

    # If root folder contains no sub-folders,
    # add root folder to LeafNode-Lst (is object-type list!).
    if($Children_ObjArr.length -eq 0){
        $Root = Get-Item -LiteralPath $SourceDir
        $LeafNode_Lst.Add($Root)
    }

    foreach($Directory in $Children_ObjArr){
        $Children_Lst.Add($Directory)
    }

    $Graph.Add($SourceDir,$Children_Lst)

    # So far we ignore loose files at root

    ### End: Insert Root Node ###

    foreach($DirectoryObject in $AllNodes_ObjectArray){

        $Children_Lst = [List[object]]::new()
        $DirectoryPath = $DirectoryObject.FullName

        # Get all direct sub-folders of any folder
        $Children_ObjArr = (Get-ChildItem -LiteralPath $DirectoryPath -Directory)

        # If folder is not a leaf node
        if($Children_ObjArr.Count -gt 0) {

            for ($i=0; $i -le ($Children_ObjArr.length - 1); $i++) {

                $null = $Children_Lst.Add($Children_ObjArr[$i])  #OBJECT!
            }

            # Add all FOLDER-Subfolder edges of FOLDER to graph.
            $null = $Graph.Add($DirectoryPath,$Children_Lst)

        }
        # If folder is a leaf-node add to LeafNode_Lst
        else{ 
            $null = $LeafNode_Lst.Add($DirectoryObject)
        }
    }

    # Add list of leaf-nodes to graph for convenience.
    $null = $Graph.Add("LeafNodes",$LeafNode_Lst)

    $AllNodes_ObjectArray = @()
}

function Get-Objects{

    Param(
        [Parameter(Mandatory)]
        [string]$SourceDir
    )

    $FoundObjectsHt = @{}
    $SkippedObjectsHt = @{}

    $ToProcessLst = [List[object]]::new()

    $VisitedNodes_HsStr = [HashSet[string]]::new()
    $Queue_Str = [Queue[string]]::new()

    # We start at root
    $null = $Queue_Str.Enqueue($SourceDir)
    $null = $VisitedNodes_HsStr.Add($SourceDir)

    $WrongExtensionCounter = 0
    $BadFolderCounter = 0

    while($Queue_Str.Count -gt 0){

        # Every directory is visited, including leaf nodes.
        $Node_Str = $Queue_Str.Dequeue() 

        ### Begin: Get loose files (Archives) ###

        # This includes loose files located at the top level of SourceDir.

        $LooseFiles_ObjArr = (Get-ChildItem -LiteralPath $Node_Str -File) # <-   -File (!)
    
        $SubDir_Count = (Get-ChildItem -Force -Directory $Node_Str).Count

        # If node is internal-node (non leaf-node) and contains files...
        if(($LooseFiles_ObjArr.Count -gt 0) -and ($SubDir_Count -gt 0)){

            foreach($LooseFile in $LooseFiles_ObjArr){

                $LooseFileExtension = [System.IO.Path]::GetExtension($LooseFile)

                # If file is archive, add to ToProcessLst.
                if( $IsArchive_Set.Contains($LooseFileExtension) ){
                    $null = $ToProcessLst.Add($LooseFile)
                }
                # If loose-file is not Archive, add to SkippedObjects.
                else{
                    $WrongExtensionCounter += 1

                    $SkippedObjectProperties = @{
                        ObjectParent = (Split-Path -Parent $LooseFile.FullName);
                        Path = $LooseFile.FullName;
                        ObjectName = $LooseFile.Name;
                        Reason = "WrongExtension";
                        Extension = "$LooseFileExtension"}

                    $SkippedObjectsHt.Add("$($LooseFile.FullName)",$SkippedObjectProperties)
                }
            }
        }

        ### End: Get loose files (Archives) ###

        # $Graph.$Node_Str <=> List of direct neighbours of Node_Str.
        for($i = 0; $i -le (($Graph.$Node_Str).Count - 1); $i ++){

            # If direct-neighbour wasn't visited before -> Enqueue
            # and add to VisitedNodes
            if(!($VisitedNodes_HsStr.Contains(($Graph.$Node_Str[$i]).FullName))){

                $Queue_Str.Enqueue(($Graph.$Node_Str[$i]).FullName)
                $null = $VisitedNodes_HsStr.Add(($Graph.$Node_Str[$i]).FullName)
            }
        }
    }

    ### Finally: Begin adding LeafNodes ###

    <#
        .NOTES
        IMPORTANT
        When encountering a leaf-folder with mixed content,
        we add the files that have an allowed extension to $ToProcessLst;
        but we do NOT add the folder itself!
    #>

    $CurrentFileExtensions_Set =  [HashSet[string]]::new()

    foreach($LeafNode in $Graph.LeafNodes){

        # Get all files in leaf-node
        $Files = (Get-ChildItem -LiteralPath $LeafNode.FullName)

        # Add all archives to $ToProcessLst
        foreach($File in $Files){

            $FileExtension = [System.IO.Path]::GetExtension($File)

            $null = $CurrentFileExtensions_Set.Add($FileExtension)

            if($IsArchive_Set.Contains($FileExtension)){
                $null = $ToProcessLst.Add($File)
            }
        }

        # Add leaf-node itself to $ToProcessLst IFF it only contains ImageFiles and is not empty.
        ## DISMISS EMPTY LEAF-NODES by ensuring that CurrentFileExtension_Set is not empty.
        if ($CurrentFileExtensions_Set.Count -gt 0){

            if ($CurrentFileExtensions_Set.isSubsetOf($IsImage_Set)) {
    
                $null = $ToProcessLst.Add($LeafNode)
    
            }
            # If folder has mixed content, add to SkippedObjects.
            elseif(!$CurrentFileExtensions_Set.isSubsetOf($IsArchive_Set)) {
    
                $BadFolderCounter += 1
    
                $SkippedObjectProperties = @{
                    ObjectParent = (Split-Path -Parent $LeafNode.FullName);
                    Path = $LeafNode.FullName;
                    ObjectName = $LeafNode.Name;
                    Reason = "UnsupportedOrMixedContent";
                    Extension = "Folder"}
    
                $SkippedObjectsHt.Add("$($LeafNode.FullName)",$SkippedObjectProperties)
            }
        }
        # Empty leaf-node
        else{

            $BadFolderCounter += 1

            $SkippedObjectProperties = @{
                ObjectParent = (Split-Path -Parent $LeafNode.FullName);
                Path = $LeafNode.FullName;
                ObjectName = $LeafNode.Name;
                Reason = "EmptyLeafNode";
                Extension = "Folder"}

            $SkippedObjectsHt.Add("$($LeafNode.FullName)",$SkippedObjectProperties)            
        }

        $CurrentFileExtensions_Set.Clear()
    }

    ### End: Add LeafNodes ###

    $FoundObjectsHt.Add("BadFolderCounter", $BadFolderCounter)
    $FoundObjectsHt.Add("WrongExtensionCounter", $WrongExtensionCounter)

    $FoundObjectsHt.Add("SkippedObjects",$SkippedObjectsHt)
    $FoundObjectsHt.Add("ToProcess",$ToProcessLst)

    $Graph.Clear()

    return $FoundObjectsHt
}

Export-ModuleMember -Function New-Graph,Get-Objects -Variable FoundObjectsHt,SkippedObjectsHt,ToProcessLst,WrongExtensionCounter,BadFolderCounter