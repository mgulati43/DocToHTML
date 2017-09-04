package docxtohtml;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.jodconverter.OfficeDocumentConverter;
import org.jodconverter.document.DocumentFormat;
import org.jodconverter.document.DocumentFormatRegistry;
import org.jodconverter.filter.DefaultFilterChain;
import org.jodconverter.filter.RefreshFilter;
import org.jodconverter.office.DefaultOfficeManagerBuilder;
import org.jodconverter.office.OfficeManager;

public class DocumentConverter {
	protected static OfficeManager officeManager;
	protected static OfficeDocumentConverter converter;
	protected static DocumentFormatRegistry formatRegistry;
	static DefaultFilterChain chain = new DefaultFilterChain(RefreshFilter.INSTANCE);
	public static void listFilesForFolder(final File folder) {
	    for (final File fileEntry : folder.listFiles()) {
	    	File outputDir = new File("html\\"+ FilenameUtils.getBaseName(fileEntry.getName()));
	        if (fileEntry.isDirectory()) {
	            listFilesForFolder(fileEntry);
	        } else {
	        	try {
					convertFilePDF(fileEntry, outputDir, chain);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
	        }
	    }
	}
	  
	public static void main(String[] args) throws Exception {
		//File inputFile = new File("docx\\sample 1.doc");
		
		final File folder = new File("ppt\\");
		
		

	
	officeManager = new DefaultOfficeManagerBuilder().build();
    converter = new OfficeDocumentConverter(officeManager);
    formatRegistry = converter.getFormatRegistry();

    officeManager.start();
    
   
    listFilesForFolder(folder);
    
    
    
    officeManager.stop();
	}
	
	protected static void convertFilePDF(final File inputFile, final File outputDir, final DefaultFilterChain chain) throws Exception {
    DocumentFormat outputFormat = formatRegistry.getFormatByExtension("html");
    File outputFile = null;
    
    if (outputDir == null) {
      outputFile = new File(FilenameUtils.getBaseName(inputFile.getName()) + "." + outputFormat.getExtension());
    } else {
      outputFile = new File(outputDir, FilenameUtils.getBaseName(inputFile.getName()) + "." + outputFormat.getExtension());
      FileUtils.deleteQuietly(outputFile);
    }
    
    converter.convert(chain, inputFile, outputFile, outputFormat);
    
    chain.reset();
  }
}