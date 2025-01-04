/*
 * abfall - convert ICS format trash calendar into a single Excel sheet
 * Copyright (C) 2011-2025 Thomas Jensen
 *
 * This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
 * License, version 3, as published by the Free Software Foundation.
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * You should have received a copy of the GNU General Public License along with this program. If not, see
 * <https://www.gnu.org/licenses/>.
 */
package com.thomasjensen.abfall;

import java.io.File;
import java.io.PrintWriter;
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
import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.io.IoBuilder;


/**
 * Command line argument parser.
 */
public class CmdLine
{
    private static final Logger LOG = LogManager.getLogger(CmdLine.class);



    private Options createOptions()
    {
        Option help = new Option("h", "help", false, "Print this message");

        Option year = Option.builder("y")
            .longOpt("year")
            .hasArg().argName("year").type(PatternOptionBuilder.NUMBER_VALUE)
            .desc("(required) The four-digit year for which the summary is being created")
            .required()
            .build();

        Option locale = Option.builder("l")
            .longOpt("locale")
            .hasArg().argName("code").type(PatternOptionBuilder.STRING_VALUE)
            .desc("Locale for printing date information (not a real translation; default: de)")
            .build();

        Option outFile = Option.builder("o")
            .longOpt("output")
            .hasArg().argName("xls").type(PatternOptionBuilder.FILE_VALUE)
            .desc("Name of the Excel file to create (output file)")
            .build();

        Options result = new Options();
        result.addOption(help);
        result.addOption(year);
        result.addOption(locale);
        result.addOption(outFile);

        return result;
    }



    private boolean hasHelp(final String[] pRawArgs)
    {
        boolean result = false;
        for (String rawArg : pRawArgs) {
            if (rawArg.equals("-h") || rawArg.equals("--help")) {
                result = true;
                break;
            }
        }
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

        if (hasHelp(pArgs)) {
            usage(opts);
            return null;
        }

        int year = -1;
        Locale locale = Locale.GERMAN;
        File inFile = null;
        File outFile = null;

        CommandLineParser parser = new DefaultParser();

        try {
            CommandLine cmd = parser.parse(opts, pArgs);

            if (cmd.hasOption('y')) {
                Long j = (Long) cmd.getParsedOptionValue("y");
                if (j != null) {
                    year = j.intValue();
                    if (year < 2000) {
                        throw new ParseException("Year too small. Must be 4 digits: " + year);
                    }
                }
            }

            if (cmd.hasOption('l')) {
                String s = cmd.getOptionValue('l');
                locale = new Locale(s);
                if (!Arrays.asList(Locale.getAvailableLocales()).contains(locale)) {
                    throw new ParseException("Unknown locale: " + s);
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
                throw new ParseException("Too many command line arguments");
            }
            if (args.size() < 1) {
                throw new ParseException("Input file not specified");
            }
            inFile = new File(args.get(0));
        }
        catch (ParseException e) {
            LOG.error(e.getMessage());
            usage(opts);
            return null;
        }
        catch (RuntimeException e) {
            LOG.error(e.getMessage(), e);
        }

        return new Config(year, locale, inFile, outFile);
    }



    private void usage(final Options pOptions)
    {
        HelpFormatter formatter = new HelpFormatter();
        final int textWidthChars = 100;
        try (PrintWriter pw = IoBuilder.forLogger(LOG).setAutoFlush(true).setLevel(Level.INFO).buildPrintWriter()) {
            formatter.printHelp(pw, textWidthChars, "abfall [options] <icsFile>", null, pOptions,
                HelpFormatter.DEFAULT_LEFT_PAD, HelpFormatter.DEFAULT_DESC_PAD, null);
        }
        LOG.info("");
    }
}
