package com.thomasjensen.abfall;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.cli.PatternOptionBuilder;


/**
 * Command line argument parser.
 */
public class CmdLine
{
    private Options createOptions()
    {
        Option help = new Option("h", "help", false, "Diesen Text anzeigen");

        Option year = Option.builder("j")
            .longOpt("jahr")
            .hasArg().argName("Jahr").type(PatternOptionBuilder.NUMBER_VALUE)
            .desc("Das vierstellige Jahr, für das der Abfallkalender erstellt wird")
            .required()
            .build();

        Option locale = Option.builder("l")
            .longOpt("locale")
            .hasArg().argName("code").type(PatternOptionBuilder.STRING_VALUE)
            .desc("Ländercode für die Datumswerte (keine vollständige Übersetzung; default: de)")
            .build();

        Option outFile = Option.builder("o")
            .longOpt("output")
            .hasArg().argName("xls").type(PatternOptionBuilder.FILE_VALUE)
            .desc("Name der erzeugten Excel-Datei (Ausgabe-Datei)")
            .build();

        Options result = new Options();
        result.addOption(help);
        result.addOption(year);
        result.addOption(locale);
        result.addOption(outFile);

        return result;
    }



    /**
     * Parse and validate the given command line arguments.
     *
     * @param pArgs actual command line arguments
     * @return config with parsed and validated arguments
     */
    public Config parse(final String[] pArgs)
    {
        Options opts = createOptions();

        int year = -1;
        Locale locale = Locale.GERMAN;
        File inFile = null;
        File outFile = null;

        CommandLineParser parser = new DefaultParser();

        try {
            CommandLine cmd = parser.parse(opts, pArgs);
            if (cmd.hasOption('h')) {
                usage(opts);
                return null;
            }

            if (cmd.hasOption('j')) {
                Long j = (Long) cmd.getParsedOptionValue("j");
                if (j != null) {
                    year = j.intValue();
                    if (year < 2000) {
                        throw new ParseException("Jahr zu klein. Muss vierstellig sein: " + year);
                    }
                }
            }

            if (cmd.hasOption('l')) {
                String s = cmd.getOptionValue('l');
                locale = new Locale(s);
                if (!Arrays.asList(Locale.getAvailableLocales()).contains(locale)) {
                    throw new ParseException("Unbekanntes locale: " + s);
                }
            }

            if (cmd.hasOption('o')) {
                outFile = new File(cmd.getOptionValue('o'));
            }
            if (outFile == null) {
                outFile = new File("Abfallkalender " + year + ".xlsx");
            }

            final List<String> args = cmd.getArgList();
            if (args.size() > 1) {
                throw new ParseException("Zu viele Argumente angegeben");
            }
            if (args.size() < 1) {
                throw new ParseException("Eingabe-Datei wurde nicht angegeben");
            }
            inFile = new File(args.get(0));
        }
        catch (ParseException e) {
            System.err.println(e.getMessage());
            usage(opts);
            return null;
        }
        catch (RuntimeException e) {
            System.err.println(e.getMessage());
            e.printStackTrace(System.err);
        }

        return new Config(year, locale, inFile, outFile);
    }



    private void usage(final Options pOptions)
    {
        HelpFormatter formatter = new HelpFormatter();
        final int textWidthChars = 100;
        formatter.setWidth(textWidthChars);
        System.out.println("Abfallkalender v1.0.0 - Abholtermine auf 1 A4-Seite zusammengefasst");  // TODO app name
        formatter.printHelp("java -jar TODO.jar [options] <icsDatei>", pOptions);  // TODO
    }
}
