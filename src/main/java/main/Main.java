package main;

import org.apache.commons.cli.*;
import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.core.LoggerContext;
import org.apache.logging.log4j.core.config.Configuration;
import org.apache.logging.log4j.core.config.LoggerConfig;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;

public class Main {
    public static final String FILE_EXTENSION = ".xlsx";
    private static Logger logger = LogManager.getLogger(Main.class);

    /**
     * Main method, it check that the input file exists and have a correct extension. Then, create
     * the output folder with the same name that the file and execute the extraction.
     * @param args Arguments come from the command line.
     */
    public static void main(String[] args) {
        CommandLine cmd = createCliOptions(args);
        File xlsx_file = new File(cmd.getOptionValue("file"));

        String extension = xlsx_file.getName().substring(xlsx_file.getName().lastIndexOf('.'));

        if (xlsx_file.exists() && extension.equalsIgnoreCase(FILE_EXTENSION)) {
            File outputFolder = createOutputDir(xlsx_file.getName());
            logger.debug(String.format("Created directory for output files in %s", outputFolder.getAbsolutePath()));
            extract(xlsx_file, outputFolder);
            logger.info(String.format("File %s founded", xlsx_file.getName()));
        } else if (!extension.equalsIgnoreCase(FILE_EXTENSION)) {
            logger.error(String.format("Input file has not a valid extension (%s)", FILE_EXTENSION));
        } else {
            logger.error(String.format("Input file %s does not exists", xlsx_file));
        }


    }

    /**
     * Read a source file (xlsx), extract each sheet in a CSV file, with the same name that it.     *
     * @param sourceFile (.xlsx) input file to extract.
     * @param ouputFolder Folder where it will be saved each CSV.
     */
    private static void extract(File sourceFile, File ouputFolder) {
        try {

            Workbook workbook = WorkbookFactory.create(sourceFile);
            logger.info(String.format("File %s has %d sheets", sourceFile.getName(), workbook.getNumberOfSheets()));

            workbook.forEach(sheet -> {
                logger.debug(String.format("Processing sheet %s of %s", sheet.getSheetName(), sheet.getSheetName()));
                String fileName = sheet.getSheetName().replaceAll("\\W+", "_") + ".csv";
                File outputFile = new File(ouputFolder, fileName);
                try (PrintWriter printWriter = new PrintWriter(new FileOutputStream(outputFile, false))) {
                    sheet.forEach(row -> {
                        String rowString = getRow(row);
                        if (!rowString.replaceAll("\\n", "").isEmpty()) {
                            printWriter.append(rowString);
                        }
                        logger.debug(String.format("Extracted from %s %5d/%d rows", sheet.getSheetName(), row.getRowNum(), sheet.getLastRowNum()));
                    });
                    logger.info(String.format("File %s has been created", fileName));
                } catch (FileNotFoundException e) {
                    logger.error("Could not be possible, create " + fileName);
                }
            });
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();

        }
    }

    /**
     * Receive a row and construct a String with all data formatted between \" and separated by ;
     *
     * @param row of sheet
     * @return Row formatted to csv
     */
    private static String getRow(Row row) {
        DataFormatter dataFormatter = new DataFormatter();

        StringBuilder row_formated = new StringBuilder();
        row.forEach(cell -> {
            String cellValue = dataFormatter.formatCellValue(cell);
            row_formated.append(String.format("\"%s\";", cellValue));
        });
        row_formated.append('\n');

        return row_formated.toString();
    }

    /**
     * Create dir to put the outputs CSV
     *
     * @param fileName Get the fileName of file to create a specific folder
     * @return File instance with the output dir path.
     */
    private static File createOutputDir(String fileName) {
        File outputDir = new File("output", fileName);
        Boolean result = outputDir.mkdirs();
        return outputDir;
    }

    /**
     * Method which create a command line parser to select input file and logger level
     *
     * @param args Args from command line
     * @return CommandLine instance to get data
     */
    private static CommandLine createCliOptions(String[] args) {
        Options options = new Options();

        Option xlsx_file = new Option("f", "file", true, "Excel file to convert to CSV");
        xlsx_file.setRequired(true);
        options.addOption(xlsx_file);

        Option verbose = new Option("v", "verbose", false, "See verbose log");
        verbose.setRequired(false);
        options.addOption(verbose);

        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();
        CommandLine cmd = null;

        try {
            cmd = parser.parse(options, args);
            changeLogLevel(cmd.hasOption("verbose") ? Level.DEBUG : Level.INFO);

        } catch (ParseException e) {
            logger.error(e);
            formatter.printHelp("Excel2CSV", options);
            System.exit(1);

        }
        return cmd;
    }

    /**
     * Method to change log level in whole project.
     * @param level new level to set the logger.
     */
    private static void changeLogLevel(Level level) {
        LoggerContext ctx = (LoggerContext) LogManager.getContext(false);
        Configuration config = ctx.getConfiguration();
        LoggerConfig loggerConfig = config.getLoggerConfig(LogManager.ROOT_LOGGER_NAME);
        loggerConfig.setLevel(level);
        ctx.updateLoggers();  // This causes all Loggers to refetch information from their LoggerConfig.
    }
}
