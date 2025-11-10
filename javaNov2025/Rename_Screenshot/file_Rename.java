package Others;

import java.io.File;

public class file_Rename {

    public static void main(String[] args) {
        // Path to the folder containing the screenshots
        String folderPath = "C:\\Users\\Keerthana.7179\\Documents\\screen";

        // Date to replace
        String oldDate = "02-08-2025";
        String newDate = "04-08-2025";

        // Get the folder
        File folder = new File(folderPath);

        // List all files in the folder
        File[] files = folder.listFiles();

        if (files != null) {
            for (File file : files) {
                String fileName = file.getName();

                // Check if the file contains the old date
                if (fileName.contains(oldDate)) {
                    // Replace the old date with the new one
                    String newFileName = fileName.replace(oldDate, newDate);

                    // Create the new file path
                    File renamedFile = new File(folderPath + File.separator + newFileName);

                    // Rename the file
                    boolean success = file.renameTo(renamedFile);

                    if (success) {
                        System.out.println("Renamed: " + fileName + " → " + newFileName);
                    } else {
                        System.out.println("Failed to rename: " + fileName);
                    }
                }
            }
        } else {
            System.out.println("Folder not found or empty.");
        }
    }
}
