import groovy.io.FileType
import org.apache.poi.poifs.filesystem.DirectoryNode
import org.apache.poi.poifs.filesystem.DocumentInputStream
import org.apache.poi.poifs.filesystem.Entry
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.poifs.filesystem.DirectoryEntry
import org.apache.poi.poifs.filesystem.DocumentEntry

println "hey"

//def input = "w:\\Code\\uft-simple-tests\\Solution1\\Test Jenkins Connection\\Parameters.mtr"
//readFile(new File(input))


def dir="w:\\\\Code\\\\uft-simple-tests\\\\Solution1\\\\Test Jenkins Connection"
new File(dir).eachFileRecurse(FileType.FILES) {
    readFile(it)
}

private void readFile(File input) {
    POIFSFileSystem fs
    try {
        fs = new POIFSFileSystem(input.newInputStream())
    }
    catch (IOException e) {
        System.err.println "Cannot read file $input.absolutePath"
        System.err.println e
        return;
    }

    println "=======Reading file: $input.absolutePath ==================="
    DirectoryEntry root = fs.getRoot()

    printDir(root)
}


private void printDir(DirectoryNode root) {
    println("DIR: $root.path $root.name $root.storageClsid")
    println("DIR desc: $root.shortDescription")


    for (Entry entry : root.entries) {
        System.out.println("found entry: " + entry.getName())
        if (entry instanceof DirectoryEntry) {
            printDir(entry)
        } else if (entry instanceof DocumentEntry) {
            printDoc(entry)
        } else {
            println("Unknown: " + entry.name)
        }
    }
}

private void printDoc(DocumentEntry entry) {
    println("Doc: $entry.name size: $entry.size")
    DocumentInputStream doc = new DocumentInputStream(entry)
    println "Content:\n $doc.text \n"
}