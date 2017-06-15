/* :name=SDLXLIFF Extract :description=Extracts internal files from sdlxliff files
 *
 * @author  Briac Pilpré
 * @date    2017-06-15
 * @version 0.2
 */
import static javax.swing.JOptionPane.*

def prop = project.projectProperties
if (!prop) {
        final def title = 'SDLXLIFF Extract';
        final def msg   = 'Please try again after you open a project.';
        showMessageDialog(null, msg, title, INFORMATION_MESSAGE);
        return;
}

outputDir = new File(prop.projectRoot, 'internal-files');
sourceDir = new File(prop.projectRoot, 'source');

outputDir.mkdirs();

class SDLXLIFFFilter implements FilenameFilter {
    public boolean accept(File f, String filename) {
        return filename.endsWith("sdlxliff") || filename.endsWith("xdliff");
    }
}

sourceDir.list(new SDLXLIFFFilter()).each { xliffFile ->
	console.println("Processing ${xliffFile}...");

	outputFileName = xliffFile.replaceAll(/\.\w+$/, '');

	def xliff = new XmlSlurper().parse(new File(sourceDir, xliffFile));
	
	def fileForm = xliff.file.header.reference['internal-file'].@form;
	if (fileForm != 'base64') {
		console.println("Internal file is not a base64 file (" + fileForm + ")");
		return;
	}

	def base64 = xliff.file.header.reference['internal-file'].text();

	def outputFile = new File(outputDir, outputFileName + ".zip");
	console.println("  -> ${outputFile} extracted.");

	outputFile.setBytes(base64.decodeBase64());

	def zipFile = new java.util.zip.ZipFile(outputFile);
	zipFile.entries().findAll { !it.directory }.each {
		def extractedFile = outputFileName + "_" + it.toString();
		console.println("    -> ${extractedFile}");
		new File(outputDir, extractedFile).text = zipFile.getInputStream(it).text
	}
	zipFile.close();
	outputFile.delete();

	console.println("\n");
}

return true;
