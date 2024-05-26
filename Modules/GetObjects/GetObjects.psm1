
using namespace System.Collections.Generic

$ErrorActionPreference = "Stop"

function New-Graph{

    Param(
        [Parameter(Mandatory)]
        [string]$SourceDir
    )

    Write-Output "Scanning folders. Please wait..."

    $Graph = @{}

    # Get all nodes (folders).
    # This does NOT include the root directory iteself.
    $AllNodes_ObjectArray = (Get-ChildItem -LiteralPath $SourceDir -Directory -Recurse)

    #
    ### [BEGIN] Insert Root Node ###
    #
    
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

    #
    ### [END] Insert Root Node ###
    #

    foreach($DirectoryObject in $AllNodes_ObjectArray){

        $Children_Lst = [List[object]]::new()
        $DirectoryPath = $DirectoryObject.FullName

        # Get all direct sub-folders of any folder
        $Children_ObjArr = (Get-ChildItem -LiteralPath $DirectoryPath -Directory)

        # If folder is not a leaf node
        if($Children_ObjArr.Count -gt 0) {

            for ($i=0; $i -le ($Children_ObjArr.length - 1); $i++) {

                $null = $Children_Lst.Add($Children_ObjArr[$i])  # OBJECT!
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

    return $Graph
}

function Get-Objects{

    Param(
        [Parameter(Mandatory)]
        [string]$SourceDir,

        [Parameter(Mandatory)]
        [hashtable]$PathsProgram,

        [Parameter(Mandatory)]
        [string]$LibraryName
    )

    $Graph = New-Graph -SourceDir $SourceDir

    
    try{
        $VisitedObjects = Import-Clixml -Path "$($PathsProgram.AppData)\VisitedObjects $LibraryName.xml"
    }
    catch{
        $VisitedObjects = @{}        
    }

    $SkippedObjects = @{}

    # CD-Folder
    # A folder that only contains elements of ComicData_Set is called a ComicData/CD-Folder
    # We include .txt in case a CD-folder contains a text-file with additional information.
    $ComicData_Set = [System.Collections.Generic.HashSet[String]] @('.jpg', '.jpeg', '.png', '.txt')
    $IsArchive_Set = [System.Collections.Generic.HashSet[String]] @('.zip', '.rar', '.cbz', '.cbr')

    $FoundObjectsHt = @{}

    $ToProcessLst = [List[object]]::new()

    $VisitedNodes_HsStr = [HashSet[string]]::new()
    $Queue_Str = [Queue[string]]::new()

    # We start at root
    $null = $Queue_Str.Enqueue($SourceDir)
    $null = $VisitedNodes_HsStr.Add($SourceDir)

    $WrongExtensionCounter = 0
    $BadFolderCounter = 0

    while($Queue_Str.Count -gt 0){

        #
        ### [BEGIN] Get loose files (Archives) ###
        #

        # Every directory is visited, including leaf nodes,
        # and loose files located at the top level of SourceDir.

        $Node_Str = $Queue_Str.Dequeue() 

        $LooseFiles_ObjArr = (Get-ChildItem -LiteralPath $Node_Str -File) # <-   -File (!)

        $SubDir_Count = (Get-ChildItem -LiteralPath $Node_Str -Directory -Force ).Count

        # If node is internal-node (non leaf-node) and contains files...
        if(($LooseFiles_ObjArr.Count -gt 0) -and ($SubDir_Count -gt 0)){

            foreach($LooseFile in $LooseFiles_ObjArr){

                if (  ! $VisitedObjects.ContainsKey($LooseFile.FullName)  ) {

                    $null = $VisitedObjects.Add("$($LooseFile.FullName)", "0")

                    $LooseFileExtension = [System.IO.Path]::GetExtension($LooseFile)

                    # If file is archive, add to ToProcessLst.
                    if ( $IsArchive_Set.Contains($LooseFileExtension) ) {
                        $null = $ToProcessLst.Add($LooseFile)
                    }
                    # If loose-file is not Archive, add to SkippedObjects.
                    else {
                        $WrongExtensionCounter += 1
                        $ObjectParent = Split-Path -Parent $LooseFile.FullName

                        $SkippedObjectProperties = @{
                            ObjectParent = $ObjectParent;
                            Path         = $LooseFile.FullName;
                            ObjectName   = $LooseFile.Name;
                            Reason       = "WrongExtension";
                            Extension    = "$LooseFileExtension"
                        }

                        if ($SkippedObjects.ContainsKey($ObjectParent)) {
                            $null = $SkippedObjects.$ObjectParent.Add("$($LooseFile.FullName)", $SkippedObjectProperties)
                        }
                        else {
                            $SkippedObjects[$ObjectParent] = @{}
                            $null = $SkippedObjects.$ObjectParent.Add("$($LooseFile.FullName)", $SkippedObjectProperties)
                        }
                        
                    }
                }
            }
        }

        #
        ### [END] Get loose files (Archives) ###
        #

        # $Graph.$Node_Str <=> List of direct neighbours of Node_Str.
        for($i = 0; $i -le (($Graph.$Node_Str).Count - 1); $i ++){

            # If direct neighbour wasn't visited before ->
            # Enqueue and add to VisitedNodes
            if(!($VisitedNodes_HsStr.Contains(($Graph.$Node_Str[$i]).FullName))){

                $Queue_Str.Enqueue(($Graph.$Node_Str[$i]).FullName)
                $null = $VisitedNodes_HsStr.Add(($Graph.$Node_Str[$i]).FullName)
            }
        }
    }

    #
    ### [BEGIN] Adding LeafNodes ###
    #

    <#
        .NOTES
        IMPORTANT
        When encountering a leaf-folder with mixed content,
        we add the files that have an allowed extension to $ToProcessLst;
        but we do NOT add the folder itself!
    #>

    $CurrentFileExtensions_Set =  [HashSet[string]]::new()

    $TemporarilyExcluded = [List[object]]::new()

    foreach($LeafNode in $Graph.LeafNodes){ # $Graph.LeafNodes is of type [List[Object]]

        if (  ! $VisitedObjects.ContainsKey($LeafNode.FullName)  ){ # < A leaf-node is added to visitedObjects IFF it only contains image-files

            # Get all files in this leaf-node
            $Files = (Get-ChildItem -LiteralPath $LeafNode.FullName)
            
            if($Files.Count -gt 0){

                foreach ($File in $Files) {
            
                    $FileExtension = [System.IO.Path]::GetExtension($File)
                    $null = $CurrentFileExtensions_Set.Add($FileExtension)
                    
                    if (! $VisitedObjects.ContainsKey($File.FullName)) {

                        if ($IsArchive_Set.Contains($FileExtension)) {
                            $null = $VisitedObjects.Add("$($File.FullName)", "0")
                            $null = $ToProcessLst.Add($File)
                        }
                        else{
                            # We do not add $File to VisitedObjects here,
                            # since we don't want to add the image-files of CD-folders.
                            $TemporarilyExcluded.Add($File)
                        }

                    }
                }

                # Add leaf-node itself to $ToProcessLst IFF it only contains ImageFiles and is not empty.
                # This only holds for CD-Folders.
                # This type of folder will _not_ be scanned for updates once discovered.
                if ($CurrentFileExtensions_Set.isSubsetOf($ComicData_Set)) {
    
                    $null = $VisitedObjects.Add("$($LeafNode.FullName)", "0")
                    $null = $ToProcessLst.Add($LeafNode)
    
                }
                else{

                    foreach($File in $TemporarilyExcluded){

                        $null = $VisitedObjects.Add("$($File.FullName)", "0")

                        # Not beautiful to do this twice, but may save us the overhead
                        # of adding an additional data-structure
                        # (or creating a bunch of hashtables that get discarded right afterwards)
                        $FileExtension = [System.IO.Path]::GetExtension($File)

                        $WrongExtensionCounter += 1
                        $ObjectParent = Split-Path -Parent $LeafNode.FullName

                        $SkippedObjectProperties = @{
                            ObjectParent = $ObjectParent;
                            Path         = $File.FullName;
                            ObjectName   = $File.Name;
                            Reason       = "WrongExtension";
                            Extension    = $FileExtension
                        }

                        if ($SkippedObjects.ContainsKey($ObjectParent)) {
                            $null = $SkippedObjects.$ObjectParent.Add("$($File.FullName)", $SkippedObjectProperties) 
                        }
                        else {
                            $SkippedObjects[$ObjectParent] = @{}
                            $null = $SkippedObjects.$ObjectParent.Add("$($File.FullName)", $SkippedObjectProperties) 
                        }
    
                                           
                    }
                }

                # A Non-CD Folder itself :
                 # Is not added to toProcessLst
                 # Is not added to VisitedObjects
                 # Is not added to SkippedObjects  

            }

            # Any empty folder will be ignored and:
             # Is not added to toProcessLst
             # Is not added to VisitedObjects
             # Is not added to SkippedObjects

        }

        $CurrentFileExtensions_Set.Clear()
        $TemporarilyExcluded.Clear()
        
    }

    #
    ### [END] Adding LeafNodes ###
    #

    $FoundObjectsHt.Add("BadFolderCounter", $BadFolderCounter)
    $FoundObjectsHt.Add("WrongExtensionCounter", $WrongExtensionCounter)

    $FoundObjectsHt.Add("SkippedObjects",$SkippedObjects)
    $FoundObjectsHt.Add("ToProcess",$ToProcessLst)

    #$Graph.Clear()

    $VisitedObjects |  Export-Clixml -Path "$($PathsProgram.AppData)\VisitedObjects $LibraryName.xml" -Force

    return $FoundObjectsHt
}

Export-ModuleMember -Function New-Graph,Get-Objects -Variable FoundObjectsHt,SkippedObjects,ToProcessLst,WrongExtensionCounter,BadFolderCounter