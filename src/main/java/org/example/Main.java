package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class Main {

    private static final double CAPTURE_OFFSET_SECONDS = 0.0;
    private static final double MIN_HIT_GAP_SECONDS = 0.5;

    private static final String FFMPEG_ENV = "FFMPEG_PATH";
    private static final String DEFAULT_SCREENSHOT_HEADING = "Test Case 1";
    private static final String DEFAULT_USER_STORY_HEADING = "General / No User Story Mentioned";

    // Detects: Test Case 1, Testcase 1, TC 1, TC1, Test Case Number 1
    private static final Pattern TEST_CASE_DIGIT_PATTERN = Pattern.compile(
            "\\b(?:test\\s*case|testcase|tc)\\s*(?:number\\s*)?(\\d+)\\b",
            Pattern.CASE_INSENSITIVE
    );

    // Detects: Test Case One, Testcase One, TC One
    private static final Pattern TEST_CASE_WORD_AFTER_PATTERN = Pattern.compile(
            "\\b(?:test\\s*case|testcase|tc)\\s*(?:number\\s*)?(" + numberWordAlternatives() + ")\\b",
            Pattern.CASE_INSENSITIVE
    );

    // Detects: First Test Case, Second Testcase, Third TC
    private static final Pattern TEST_CASE_WORD_BEFORE_PATTERN = Pattern.compile(
            "\\b(" + numberWordAlternatives() + ")\\s+(?:test\\s*case|testcase|tc)\\b",
            Pattern.CASE_INSENSITIVE
    );

    // Detects: Scenario 1, Scenario One
    private static final Pattern SCENARIO_DIGIT_PATTERN = Pattern.compile(
            "\\bscenario\\s*(\\d+)\\b",
            Pattern.CASE_INSENSITIVE
    );

    private static final Pattern SCENARIO_WORD_PATTERN = Pattern.compile(
            "\\bscenario\\s*(" + numberWordAlternatives() + ")\\b",
            Pattern.CASE_INSENSITIVE
    );

    // Detects:
    // User Story 456789
    // Story 456789
    // Story Number 456789
    // Story number is 456789
    // US 456789
    // ADO Story 456789
    private static final Pattern USER_STORY_PATTERN = Pattern.compile(
            "\\b(?:user\\s*story|userstory|ado\\s*story|story|us)" +
                    "\\s*(?:number|no|id)?\\s*(?:is)?\\s*#?\\s*(\\d{4,})\\b",
            Pattern.CASE_INSENSITIVE
    );

    // Detects: Bug, Bug 12345, Defect, Defect 12345, Issue, Issue 12345
    private static final Pattern BUG_PATTERN = Pattern.compile(
            "\\b(?:bug|defect|issue)\\s*#?\\s*(\\d+)?\\b",
            Pattern.CASE_INSENSITIVE
    );

    // Detects: next test case, next testcase, next test, moving to next test, go to next test
    private static final Pattern NEXT_TEST_CASE_PATTERN = Pattern.compile(
            "\\b(?:next|moving\\s+to\\s+next|go\\s+to\\s+next|start\\s+next|another)\\s+(?:test\\s*case|testcase|test|tc)\\b",
            Pattern.CASE_INSENSITIVE
    );

    private static final List<String> SCREENSHOT_TRIGGERS = List.of(
            "screenshot",
            "screenshots",
            "screen shot",
            "screen-shot",
            "screen shots",

            "pin shot",
            "pin short",
            "pin chart",
            "pinchart",

            "in shot",
            "in short",
            "in chart",
            "inchart",

            "green shot",
            "green short",
            "green chart",

            "screen chart",
            "screen short",

            "scream shot",
            "scream short",
            "scream chart",

            "spring shot",
            "spring short",
            "spring chart",

            "print shot",
            "print short",
            "print chart",

            "skin shot",
            "skin short",
            "skin chart",

            "clean shot",
            "clean short",
            "clean chart",

            "stream shot",
            "stream short",
            "stream chart",

            "beingshot"
    );

    private static final Pattern TEAMS_WORD_TIME_PATTERN = Pattern.compile(
            "(?:(\\d+)\\s+hours?[,\\s]+)?" +
                    "(?:(\\d+)\\s+minutes?[,\\s]+)?" +
                    "(\\d+)\\s+seconds?",
            Pattern.CASE_INSENSITIVE
    );

    private static final Pattern NUMERIC_TIME_PATTERN = Pattern.compile(
            "(?<!\\d)(\\d{1,2}):(\\d{2})(?::(\\d{2}))?(?:[\\.,](\\d{1,3}))?(?!\\d)"
    );

    private static final Pattern VTT_TIMESTAMP_LINE_PATTERN = Pattern.compile(
            "^\\s*(\\d{1,2}:\\d{2}:\\d{2}[\\.,]\\d{1,3}|\\d{1,2}:\\d{2}[\\.,]\\d{1,3})\\s*-->\\s*" +
                    "(\\d{1,2}:\\d{2}:\\d{2}[\\.,]\\d{1,3}|\\d{1,2}:\\d{2}[\\.,]\\d{1,3}).*$",
            Pattern.CASE_INSENSITIVE
    );

    private static final Pattern VTT_SPEAKER_PATTERN =
            Pattern.compile("^\\s*<v\\s+([^>]+)>(.*)$", Pattern.CASE_INSENSITIVE);

    public static void main(String[] args) throws Exception {
        if (args.length < 3) {
            System.err.println("""
                    Usage:
                      java -jar AIScreenshot.jar <video-file> <teams-transcript-file> <output-docx>

                    Examples:
                      java -jar AIScreenshot.jar recording.mp4 transcript.txt screenshots.docx
                      java -jar AIScreenshot.jar recording.mp4 transcript.vtt screenshots.docx
                      java -jar AIScreenshot.jar recording.mp4 transcript.docx screenshots.docx
                    """);
            System.exit(1);
        }

        Path videoFile = Paths.get(args[0]);
        Path transcriptFile = Paths.get(args[1]);
        Path outputDocx = Paths.get(args[2]);

        if (!Files.exists(videoFile)) {
            throw new IllegalArgumentException("Video file not found: " + videoFile);
        }

        if (!Files.exists(transcriptFile)) {
            throw new IllegalArgumentException("Transcript file not found: " + transcriptFile);
        }

        if (transcriptFile.toAbsolutePath().normalize().equals(outputDocx.toAbsolutePath().normalize())) {
            throw new IllegalArgumentException("Transcript file and output DOCX cannot be the same file.");
        }

        Path tempDir = Files.createTempDirectory("teams-transcript-screenshots-");

        try {
            ProcessingResult result = processTranscript(videoFile, transcriptFile, outputDocx, tempDir);

            try (XWPFDocument document = new XWPFDocument()) {
                addQAEvidenceTitle(document);

                if (result.hits().isEmpty()) {
                    addParagraph(document, "No screenshot triggers found in the Teams transcript.");
                } else {
                    addTableOfContents(document, result.hits());
                    addPageBreak(document);
                    addGroupedScreenshotEntries(document, result);
                }

                //Saved Word document to     outputDocx
                try (OutputStream outputStream = Files.newOutputStream(outputDocx)) {
                    document.write(outputStream);
                }
            }

        } finally {
            deleteRecursively(tempDir);
        }
    }

    private static ProcessingResult processTranscript(
            Path videoFile,
            Path transcriptFile,
            Path outputDocx,
            Path tempDir
    ) throws IOException, InterruptedException {

        List<TranscriptEntry> entries = readTeamsTranscript(transcriptFile, outputDocx);
        List<TranscriptHit> hits = new ArrayList<>();
        List<SectionTranscriptLine> sectionTranscriptLines = new ArrayList<>();
        String currentUserStoryHeading = DEFAULT_USER_STORY_HEADING;
        int currentTestCaseNumber = 1;
        String currentTestCaseHeading = DEFAULT_SCREENSHOT_HEADING;

        for (TranscriptEntry entry : entries) {
            Optional<String> userStoryHeading = findUserStoryHeading(entry.text());

            if (userStoryHeading.isPresent()) {
                currentUserStoryHeading = userStoryHeading.get();
                currentTestCaseNumber = 1;
                currentTestCaseHeading = DEFAULT_SCREENSHOT_HEADING;
            }

            Optional<TestCaseMarker> testCaseMarker = findTestCaseMarker(entry.text(), currentTestCaseNumber);

            if (testCaseMarker.isPresent()) {
                currentTestCaseNumber = testCaseMarker.get().number();
                currentTestCaseHeading = testCaseMarker.get().heading();
            }

            Optional<String> bugHeading = findBugHeading(entry.text());

            if (bugHeading.isPresent()) {
                currentTestCaseHeading = bugHeading.get();
            }

            sectionTranscriptLines.add(new SectionTranscriptLine(
                    currentUserStoryHeading,
                    currentTestCaseHeading,
                    entry.seconds(),
                    entry.speaker(),
                    entry.text()
            ));

            MatchResult matchResult = findMatchedTrigger(entry.text());


            if (matchResult.matched()) {
                double captureSeconds = Math.max(0, entry.seconds() + CAPTURE_OFFSET_SECONDS);

                if (shouldRecordHit(hits, captureSeconds)) {
                    Path screenshotPath = tempDir.resolve(
                            String.format(Locale.ROOT, "screenshot-%03d.png", hits.size() + 1)
                    );

                    captureFrameAtTime(videoFile, captureSeconds, screenshotPath);

                    hits.add(new TranscriptHit(
                            entry.seconds(),
                            captureSeconds,
                            entry.speaker(),
                            entry.text(),
                            matchResult.trigger(),
                            currentUserStoryHeading,
                            currentTestCaseHeading,
                            screenshotPath
                    ));
                }
            }
        }

        return new ProcessingResult(hits, entries, sectionTranscriptLines);
    }

    private static List<TranscriptEntry> readTeamsTranscript(Path transcriptFile, Path outputDocx) throws IOException {
        String fileName = transcriptFile.getFileName().toString().toLowerCase(Locale.ROOT);

        String content;

        if (fileName.endsWith(".docx")) {
            content = readDocxAsText(transcriptFile);
        } else {
            content = Files.readString(transcriptFile, StandardCharsets.UTF_8);
        }

        content = removeBom(content);
        //Extracted transcript text saved to debug file
        Path debugFile = createDebugTranscriptPath(outputDocx);
        // Files.writeString(debugFile, content, StandardCharsets.UTF_8);

        if (isVttTranscript(content, transcriptFile)) {
            return readVttTranscript(content);
        }

        return readPlainTeamsTranscript(content);
    }

    private static Path createDebugTranscriptPath(Path outputDocx) {
        Path parent = outputDocx.toAbsolutePath().getParent();

        if (parent == null) {
            parent = Paths.get(".").toAbsolutePath();
        }

        String outputName = outputDocx.getFileName().toString();

        return parent.resolve(outputName + "-extracted-transcript.txt");
    }

    private static String readDocxAsText(Path docxFile) throws IOException {
        StringBuilder text = new StringBuilder();

        try (InputStream inputStream = Files.newInputStream(docxFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {

            for (IBodyElement element : document.getBodyElements()) {
                if (element.getElementType() == BodyElementType.PARAGRAPH) {
                    XWPFParagraph paragraph = (XWPFParagraph) element;
                    appendLine(text, paragraph.getText());
                } else if (element.getElementType() == BodyElementType.TABLE) {
                    XWPFTable table = (XWPFTable) element;

                    for (XWPFTableRow row : table.getRows()) {
                        for (XWPFTableCell cell : row.getTableCells()) {
                            for (XWPFParagraph paragraph : cell.getParagraphs()) {
                                appendLine(text, paragraph.getText());
                            }
                        }
                    }
                }
            }
        }

        return text.toString();
    }

    private static void appendLine(StringBuilder builder, String line) {
        if (line == null) {
            return;
        }

        String clean = line.trim();

        if (!clean.isBlank()) {
            builder.append(clean).append(System.lineSeparator());
        }
    }

    private static boolean isVttTranscript(String content, Path transcriptFile) {
        String fileName = transcriptFile.getFileName().toString().toLowerCase(Locale.ROOT);

        return fileName.endsWith(".vtt")
                || content.trim().startsWith("WEBVTT")
                || content.contains("-->");
    }

    private static List<TranscriptEntry> readPlainTeamsTranscript(String content) {
        List<String> lines = Arrays.asList(content.split("\\R"));
        List<TranscriptEntry> entries = new ArrayList<>();

        List<String> textBuffer = new ArrayList<>();

        for (String rawLine : lines) {
            String line = removeBom(rawLine).trim();

            if (line.isBlank()) {
                continue;
            }

            Optional<TimestampInfo> timestampInfo = parseAnyTimestampLine(line);

            if (timestampInfo.isPresent()) {
                TimestampInfo info = timestampInfo.get();

                String spokenText;

                if (!textBuffer.isEmpty()) {
                    spokenText = String.join(" ", textBuffer).trim();
                } else {
                    spokenText = info.textBeforeTimestamp().trim();
                }

                if (spokenText.isBlank()) {
                    spokenText = line;
                }

                entries.add(new TranscriptEntry(
                        info.seconds(),
                        info.speaker(),
                        spokenText
                ));

                textBuffer.clear();
            } else {
                textBuffer.add(line);
            }
        }

        return entries;
    }

    private static Optional<TimestampInfo> parseAnyTimestampLine(String line) {
        Optional<TimestampInfo> wordTime = parseWordTimestampLine(line);

        if (wordTime.isPresent()) {
            return wordTime;
        }

        return parseNumericTimestampLine(line);
    }

    private static Optional<TimestampInfo> parseWordTimestampLine(String line) {
        Matcher matcher = TEAMS_WORD_TIME_PATTERN.matcher(line);

        TimestampInfo lastMatch = null;

        while (matcher.find()) {
            int hours = parseIntOrZero(matcher.group(1));
            int minutes = parseIntOrZero(matcher.group(2));
            int seconds = parseIntOrZero(matcher.group(3));

            double totalSeconds = hours * 3600.0 + minutes * 60.0 + seconds;

            String beforeTimestamp = line.substring(0, matcher.start()).trim();
            String speaker = beforeTimestamp.isBlank() ? "Unknown Speaker" : beforeTimestamp;

            lastMatch = new TimestampInfo(totalSeconds, speaker, beforeTimestamp);
        }

        return Optional.ofNullable(lastMatch);
    }

    private static Optional<TimestampInfo> parseNumericTimestampLine(String line) {
        Matcher matcher = NUMERIC_TIME_PATTERN.matcher(line);

        TimestampInfo lastMatch = null;

        while (matcher.find()) {
            String first = matcher.group(1);
            String second = matcher.group(2);
            String third = matcher.group(3);
            String millis = matcher.group(4);

            double totalSeconds;

            if (third != null) {
                int hours = Integer.parseInt(first);
                int minutes = Integer.parseInt(second);
                double seconds = parseSecondsWithMillis(third, millis);
                totalSeconds = hours * 3600.0 + minutes * 60.0 + seconds;
            } else {
                int minutes = Integer.parseInt(first);
                double seconds = parseSecondsWithMillis(second, millis);
                totalSeconds = minutes * 60.0 + seconds;
            }

            String beforeTimestamp = line.substring(0, matcher.start()).trim();
            String speaker = beforeTimestamp.isBlank() ? "Unknown Speaker" : beforeTimestamp;

            lastMatch = new TimestampInfo(totalSeconds, speaker, beforeTimestamp);
        }

        return Optional.ofNullable(lastMatch);
    }

    private static double parseSecondsWithMillis(String secondsText, String millisText) {
        double seconds = Double.parseDouble(secondsText);

        if (millisText != null && !millisText.isBlank()) {
            String padded = (millisText + "000").substring(0, 3);
            seconds += Integer.parseInt(padded) / 1000.0;
        }

        return seconds;
    }

    private static List<TranscriptEntry> readVttTranscript(String content) {
        List<String> lines = Arrays.asList(content.split("\\R"));
        List<TranscriptEntry> entries = new ArrayList<>();

        int i = 0;

        while (i < lines.size()) {
            String line = removeBom(lines.get(i)).trim();

            Matcher timestampMatcher = VTT_TIMESTAMP_LINE_PATTERN.matcher(line);

            if (!timestampMatcher.matches()) {
                i++;
                continue;
            }

            double startSeconds = parseVttTimeToSeconds(timestampMatcher.group(1));
            i++;

            List<String> cueTextLines = new ArrayList<>();
            String speaker = "Unknown Speaker";

            while (i < lines.size()) {
                String cueLine = removeBom(lines.get(i)).trim();

                if (cueLine.isBlank()) {
                    break;
                }

                Matcher speakerMatcher = VTT_SPEAKER_PATTERN.matcher(cueLine);

                if (speakerMatcher.matches()) {
                    speaker = speakerMatcher.group(1).trim();

                    String textPart = removeVttTags(speakerMatcher.group(2).trim());

                    if (!textPart.isBlank()) {
                        cueTextLines.add(textPart);
                    }
                } else {
                    String cleanLine = removeVttTags(cueLine);

                    if (!cleanLine.isBlank()) {
                        cueTextLines.add(cleanLine);
                    }
                }

                i++;
            }

            String spokenText = String.join(" ", cueTextLines)
                    .replaceAll("\\s+", " ")
                    .trim();

            entries.add(new TranscriptEntry(
                    startSeconds,
                    speaker,
                    spokenText
            ));

            i++;
        }

        return entries;
    }

    private static double parseVttTimeToSeconds(String value) {
        String clean = value.trim().replace(',', '.');
        String[] parts = clean.split(":");

        if (parts.length == 3) {
            int hours = Integer.parseInt(parts[0]);
            int minutes = Integer.parseInt(parts[1]);
            double seconds = Double.parseDouble(parts[2]);

            return hours * 3600.0 + minutes * 60.0 + seconds;
        }

        if (parts.length == 2) {
            int minutes = Integer.parseInt(parts[0]);
            double seconds = Double.parseDouble(parts[1]);

            return minutes * 60.0 + seconds;
        }

        throw new IllegalArgumentException("Invalid VTT timestamp: " + value);
    }

    private static MatchResult findMatchedTrigger(String text) {
        String normalized = normalizeTriggerText(text);
        String compact = normalized.replace(" ", "");

        if (normalized.isBlank()) {
            return new MatchResult(false, "NO MATCH");
        }

        for (String trigger : SCREENSHOT_TRIGGERS) {
            String triggerNormalized = normalizeTriggerText(trigger);
            String triggerCompact = triggerNormalized.replace(" ", "");

            if (normalized.contains(triggerNormalized) || compact.contains(triggerCompact)) {
                return new MatchResult(true, trigger);
            }
        }

        return new MatchResult(false, "NO MATCH");
    }

    static boolean containsScreenshot(String text) {
        return findMatchedTrigger(text).matched();
    }

    static Optional<String> detectTestCaseHeading(String text, int currentTestCaseNumber) {
        return findTestCaseMarker(text, currentTestCaseNumber).map(TestCaseMarker::heading);
    }

    private static Optional<TestCaseMarker> findTestCaseMarker(String text, int currentTestCaseNumber) {
        String normalized = normalizeTriggerText(text);

        if (normalized.isBlank()) {
            return Optional.empty();
        }

        Optional<TestCaseMarker> scenarioMarker = findScenarioMarker(normalized);

        if (scenarioMarker.isPresent()) {
            return scenarioMarker;
        }

        OptionalInt explicitNumber = findExplicitTestCaseNumber(normalized);

        if (explicitNumber.isPresent()) {
            int number = explicitNumber.getAsInt();
            return Optional.of(new TestCaseMarker(number, "Test Case " + number));
        }

        if (NEXT_TEST_CASE_PATTERN.matcher(normalized).find()) {
            int nextNumber = Math.max(1, currentTestCaseNumber + 1);
            return Optional.of(new TestCaseMarker(nextNumber, "Test Case " + nextNumber));
        }

        return Optional.empty();
    }

    private static Optional<String> findUserStoryHeading(String text) {
        String normalizedText = normalizeTriggerText(text);

        if (normalizedText.isBlank()) {
            return Optional.empty();
        }

        Matcher userStoryMatcher = USER_STORY_PATTERN.matcher(normalizedText);

        if (userStoryMatcher.find()) {
            return Optional.of("User Story " + userStoryMatcher.group(1));
        }

        return Optional.empty();
    }

    private static Optional<String> findBugHeading(String text) {
        String normalizedText = normalizeTriggerText(text);

        if (normalizedText.isBlank()) {
            return Optional.empty();
        }

        Matcher bugMatcher = BUG_PATTERN.matcher(normalizedText);

        if (bugMatcher.find()) {
            String phrase = bugMatcher.group(0).trim();
            String number = bugMatcher.group(1);

            if (phrase.startsWith("defect")) {
                return Optional.of(number == null || number.isBlank() ? "Defect" : "Defect " + number);
            }

            if (phrase.startsWith("issue")) {
                return Optional.of(number == null || number.isBlank() ? "Issue" : "Issue " + number);
            }

            return Optional.of(number == null || number.isBlank() ? "Bug" : "Bug " + number);
        }

        return Optional.empty();
    }

    private static Optional<TestCaseMarker> findScenarioMarker(String normalizedText) {
        Matcher scenarioDigitMatcher = SCENARIO_DIGIT_PATTERN.matcher(normalizedText);

        if (scenarioDigitMatcher.find()) {
            int number = Integer.parseInt(scenarioDigitMatcher.group(1));
            return Optional.of(new TestCaseMarker(number, "Scenario " + number));
        }

        Matcher scenarioWordMatcher = SCENARIO_WORD_PATTERN.matcher(normalizedText);

        if (scenarioWordMatcher.find()) {
            OptionalInt number = parseNumberWord(scenarioWordMatcher.group(1));

            if (number.isPresent()) {
                return Optional.of(new TestCaseMarker(number.getAsInt(), "Scenario " + number.getAsInt()));
            }
        }

        return Optional.empty();
    }

    private static OptionalInt findExplicitTestCaseNumber(String normalizedText) {
        Matcher digitMatcher = TEST_CASE_DIGIT_PATTERN.matcher(normalizedText);

        if (digitMatcher.find()) {
            return OptionalInt.of(Integer.parseInt(digitMatcher.group(1)));
        }

        Matcher wordAfterMatcher = TEST_CASE_WORD_AFTER_PATTERN.matcher(normalizedText);

        if (wordAfterMatcher.find()) {
            return parseNumberWord(wordAfterMatcher.group(1));
        }

        Matcher wordBeforeMatcher = TEST_CASE_WORD_BEFORE_PATTERN.matcher(normalizedText);

        if (wordBeforeMatcher.find()) {
            return parseNumberWord(wordBeforeMatcher.group(1));
        }

        return OptionalInt.empty();
    }

    private static OptionalInt parseNumberWord(String value) {
        if (value == null) {
            return OptionalInt.empty();
        }

        return switch (value.toLowerCase(Locale.ROOT)) {
            case "one", "first" -> OptionalInt.of(1);
            case "two", "second" -> OptionalInt.of(2);
            case "three", "third" -> OptionalInt.of(3);
            case "four", "fourth" -> OptionalInt.of(4);
            case "five", "fifth" -> OptionalInt.of(5);
            case "six", "sixth" -> OptionalInt.of(6);
            case "seven", "seventh" -> OptionalInt.of(7);
            case "eight", "eighth" -> OptionalInt.of(8);
            case "nine", "ninth" -> OptionalInt.of(9);
            case "ten", "tenth" -> OptionalInt.of(10);
            case "eleven", "eleventh" -> OptionalInt.of(11);
            case "twelve", "twelfth" -> OptionalInt.of(12);
            case "thirteen", "thirteenth" -> OptionalInt.of(13);
            case "fourteen", "fourteenth" -> OptionalInt.of(14);
            case "fifteen", "fifteenth" -> OptionalInt.of(15);
            case "sixteen", "sixteenth" -> OptionalInt.of(16);
            case "seventeen", "seventeenth" -> OptionalInt.of(17);
            case "eighteen", "eighteenth" -> OptionalInt.of(18);
            case "nineteen", "nineteenth" -> OptionalInt.of(19);
            case "twenty", "twentieth" -> OptionalInt.of(20);
            default -> OptionalInt.empty();
        };
    }

    private static String numberWordAlternatives() {
        return "one|first|two|second|three|third|four|fourth|five|fifth|six|sixth|seven|seventh|"
                + "eight|eighth|nine|ninth|ten|tenth|eleven|eleventh|twelve|twelfth|"
                + "thirteen|thirteenth|fourteen|fourteenth|fifteen|fifteenth|"
                + "sixteen|sixteenth|seventeen|seventeenth|eighteen|eighteenth|"
                + "nineteen|nineteenth|twenty|twentieth";
    }

    private static String normalizeTriggerText(String text) {
        if (text == null) {
            return "";
        }

        return text.toLowerCase(Locale.ROOT)
                .replaceAll("[^a-z0-9]+", " ")
                .replaceAll("\\s+", " ")
                .trim();
    }

    private static boolean shouldRecordHit(List<TranscriptHit> hits, double captureSeconds) {
        if (hits.isEmpty()) {
            return true;
        }

        TranscriptHit lastHit = hits.get(hits.size() - 1);
        return captureSeconds - lastHit.captureSeconds() >= MIN_HIT_GAP_SECONDS;
    }

    private static void captureFrameAtTime(
            Path videoFile,
            double seconds,
            Path outputImage
    ) throws IOException, InterruptedException {

        Process process = new ProcessBuilder(
                resolveExecutable(FFMPEG_ENV, "ffmpeg"),
                "-y",
                "-i",
                videoFile.toString(),
                "-ss",
                String.format(Locale.ROOT, "%.3f", seconds),
                "-frames:v",
                "1",
                "-q:v",
                "2",
                outputImage.toString()
        )
                .redirectErrorStream(true)
                .start();

        String output = readProcessOutput(process);
        int exitCode = process.waitFor();

        if (exitCode != 0) {
            throw new IOException("ffmpeg screenshot capture failed at "
                    + formatSeconds(seconds) + ":\n" + output);
        }
    }

    private static String readProcessOutput(Process process) throws IOException {
        StringBuilder output = new StringBuilder();

        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(process.getInputStream(), StandardCharsets.UTF_8))) {

            String line;

            while ((line = reader.readLine()) != null) {
                output.append(line).append('\n');
            }
        }

        return output.toString();
    }

    private static void addTitle(XWPFDocument document, String title) {
        XWPFParagraph paragraph = document.createParagraph();

        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setFontSize(16);
        run.setText(title);
    }

    /*private static void addSummary(XWPFDocument document, ProcessingResult result) {
        addParagraph(document, "Total transcript entries found: " + result.entries().size());
        addParagraph(document, "Total screenshots captured: " + result.hits().size());
        addParagraph(document, "Capture offset seconds: " + CAPTURE_OFFSET_SECONDS);
        addParagraph(document, "Minimum hit gap seconds: " + MIN_HIT_GAP_SECONDS);
    }*/

    private static void addParagraph(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();

        XWPFRun run = paragraph.createRun();
        run.setText(text == null ? "" : text);
    }

    private static void addScreenshotEntry(
            XWPFDocument document,
            int index,
            TranscriptHit hit,
            Path imagePath
    ) throws IOException, InvalidFormatException {

        XWPFParagraph heading = document.createParagraph();
        XWPFRun headingRun = heading.createRun();

        headingRun.setBold(true);
        headingRun.setFontSize(12);
        headingRun.setText("Screenshot " + index + " (" + formatSeconds(hit.captureSeconds()) + ")");

//        addParagraph(document, "Transcript Time: " + formatSeconds(hit.transcriptSeconds()));
//        addParagraph(document, "Captured Video Time: " + formatSeconds(hit.captureSeconds()));
//        addParagraph(document, "Speaker: " + hit.speaker());
//        addParagraph(document, "Matched Trigger: " + hit.matchedTrigger());
//        addParagraph(document, "Transcript Text: " + hit.text());

        XWPFParagraph imageParagraph = document.createParagraph();
        XWPFRun imageRun = imageParagraph.createRun();

        try (InputStream inputStream = Files.newInputStream(imagePath)) {
            imageRun.addPicture(
                    inputStream,
                    XWPFDocument.PICTURE_TYPE_PNG,
                    imagePath.getFileName().toString(),
                    Units.toEMU(520),
                    Units.toEMU(292)
            );
        }
    }

    private static void addGroupedScreenshotEntries(
            XWPFDocument document,
            ProcessingResult result
    ) throws IOException, InvalidFormatException {

        List<TranscriptHit> hits = result.hits();
        String previousUserStory = null;
        String previousSection = null;

        for (int index = 0; index < hits.size(); index++) {
            TranscriptHit hit = hits.get(index);

            String userStoryHeading = safeHeading(hit.userStoryHeading(), DEFAULT_USER_STORY_HEADING);
            String sectionHeading = safeHeading(hit.testCaseHeading(), DEFAULT_SCREENSHOT_HEADING);

            if (!userStoryHeading.equals(previousUserStory)) {
                if (previousUserStory != null) {
                    addMajorDivider(document);
                }

                addUserStoryHeading(document, userStoryHeading);
                previousUserStory = userStoryHeading;
                previousSection = null;
            }

            if (!sectionHeading.equals(previousSection)) {
                if (previousSection != null) {
                    addMinorDivider(document);
                }

                addTestCaseHeading(document, sectionHeading);
                previousSection = sectionHeading;
            }

            addScreenshotEntry(document, index + 1, hit, hit.imagePath());
        }
    }

    private static void addQAEvidenceTitle(XWPFDocument document) {
        XWPFParagraph title = document.createParagraph();
        XWPFRun titleRun = title.createRun();
        titleRun.setBold(true);
        titleRun.setFontSize(20);
        titleRun.setText("QA Evidence");

        XWPFParagraph generated = document.createParagraph();
        XWPFRun generatedRun = generated.createRun();
        generatedRun.setFontSize(10);
        generatedRun.setText("Generated On: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("MM/dd/yyyy hh:mm a")));
    }

    private static void addTableOfContents(XWPFDocument document, List<TranscriptHit> hits) {
        addMajorDivider(document);

        XWPFParagraph heading = document.createParagraph();
        XWPFRun headingRun = heading.createRun();
        headingRun.setBold(true);
        headingRun.setFontSize(14);
        headingRun.setText("TABLE OF CONTENTS");

        Map<String, LinkedHashMap<String, Integer>> counts = buildTableOfContentsCounts(hits);

        for (Map.Entry<String, LinkedHashMap<String, Integer>> storyEntry : counts.entrySet()) {
            XWPFParagraph storyParagraph = document.createParagraph();
            XWPFRun storyRun = storyParagraph.createRun();
            storyRun.setBold(true);
            storyRun.setFontSize(12);
            storyRun.setText(storyEntry.getKey());

            for (Map.Entry<String, Integer> sectionEntry : storyEntry.getValue().entrySet()) {
                XWPFParagraph sectionParagraph = document.createParagraph();
                sectionParagraph.setIndentationLeft(360);
                XWPFRun sectionRun = sectionParagraph.createRun();
                sectionRun.setFontSize(11);
                sectionRun.setText("• " + sectionEntry.getKey() + " (" + sectionEntry.getValue() + " " + pluralizeScreenshot(sectionEntry.getValue()) + ")");
            }
        }

        addMajorDivider(document);
    }

    private static Map<String, LinkedHashMap<String, Integer>> buildTableOfContentsCounts(List<TranscriptHit> hits) {
        Map<String, LinkedHashMap<String, Integer>> counts = new LinkedHashMap<>();

        for (TranscriptHit hit : hits) {
            String userStory = safeHeading(hit.userStoryHeading(), DEFAULT_USER_STORY_HEADING);
            String section = safeHeading(hit.testCaseHeading(), DEFAULT_SCREENSHOT_HEADING);

            counts.computeIfAbsent(userStory, key -> new LinkedHashMap<>())
                    .merge(section, 1, Integer::sum);
        }

        return counts;
    }

    private static String pluralizeScreenshot(int count) {
        return count == 1 ? "screenshot" : "screenshots";
    }

    private static void addPageBreak(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setPageBreak(true);
    }

    private static void addMajorDivider(XWPFDocument document) {
        addParagraph(document, "=========================================================");
    }

    private static void addMinorDivider(XWPFDocument document) {
        addParagraph(document, "---------------------------------------------------------");
    }

    private static String safeHeading(String value, String defaultValue) {
        if (value == null || value.isBlank()) {
            return defaultValue;
        }

        return value;
    }

    private static Map<String, List<SectionTranscriptLine>> groupLinesByUserStory(List<SectionTranscriptLine> lines) {
        Map<String, List<SectionTranscriptLine>> grouped = new LinkedHashMap<>();

        for (SectionTranscriptLine line : lines) {
            String userStory = safeHeading(line.userStoryHeading(), DEFAULT_USER_STORY_HEADING);
            grouped.computeIfAbsent(userStory, key -> new ArrayList<>()).add(line);
        }

        return grouped;
    }

    private static Map<String, List<SectionTranscriptLine>> groupLinesByUserStoryAndTestCase(List<SectionTranscriptLine> lines) {
        Map<String, List<SectionTranscriptLine>> grouped = new LinkedHashMap<>();

        for (SectionTranscriptLine line : lines) {
            String userStory = safeHeading(line.userStoryHeading(), DEFAULT_USER_STORY_HEADING);
            String testCase = safeHeading(line.testCaseHeading(), DEFAULT_SCREENSHOT_HEADING);
            grouped.computeIfAbsent(sectionKey(userStory, testCase), key -> new ArrayList<>()).add(line);
        }

        return grouped;
    }

    private static String sectionKey(String userStoryHeading, String testCaseHeading) {
        return userStoryHeading + " || " + testCaseHeading;
    }

    private static void addSummaryBlock(XWPFDocument document, String title, SummaryInfo summaryInfo) {
        XWPFParagraph titleParagraph = document.createParagraph();
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setBold(true);
        titleRun.setFontSize(11);
        titleRun.setText(title);

        addParagraph(document, "Status: " + summaryInfo.status());

        if (summaryInfo.summaryLines().isEmpty()) {
            addParagraph(document, "Summary: No clear discussion summary was detected from the transcript for this section.");
        } else {
            addParagraph(document, "Summary:");

            for (String line : summaryInfo.summaryLines()) {
                addParagraph(document, "• " + line);
            }
        }
    }

    private static SummaryInfo summarizeLines(List<SectionTranscriptLine> lines) {
        if (lines == null || lines.isEmpty()) {
            return new SummaryInfo("Not Determined", List.of());
        }

        List<String> cleanedLines = new ArrayList<>();

        for (SectionTranscriptLine line : lines) {
            String cleaned = cleanTranscriptLineForSummary(line.text());

            if (!cleaned.isBlank()) {
                cleanedLines.add(cleaned);
            }
        }

        String combined = String.join(". ", cleanedLines)
                .replaceAll("\\s+", " ")
                .trim();

        String status = inferStatus(combined);
        List<String> summaryLines = extractSummaryLines(cleanedLines);

        return new SummaryInfo(status, summaryLines);
    }

    private static String cleanTranscriptLineForSummary(String text) {
        if (text == null) {
            return "";
        }

        String cleaned = removeVttTags(text)
                .replaceAll("(?i)\\b(user\\s*story|us)\\s*#?\\s*\\d{4,}\\b", " ")
                .replaceAll("(?i)\\b(?:test\\s*case|testcase|tc)\\s*(?:number\\s*)?\\d+\\b", " ")
                .replaceAll("(?i)\\b(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)\\s+(?:test\\s*case|testcase|tc)\\b", " ")
                .replaceAll("(?i)\\bnext\\s+(?:test\\s*case|testcase|test|tc)\\b", " ")
                .replaceAll("(?i)\\bscreen\\s*shot[s]?\\b|\\bscreenshot[s]?\\b", " ")
                .replaceAll("\\s+", " ")
                .trim();

        if (cleaned.length() < 4) {
            return "";
        }

        return cleaned;
    }

    private static String inferStatus(String combinedText) {
        String normalized = normalizeTriggerText(combinedText);

        if (normalized.isBlank()) {
            return "Not Determined";
        }

        if (containsAny(normalized,
                "blocked", "unable", "cannot", "can not", "access issue", "permission issue")) {
            return "Blocked / Needs Assistance";
        }

        if (containsAny(normalized,
                "not working", "failed", "failure", "error", "issue", "defect", "bug", "incorrect", "mismatch", "did not", "doesn t", "does not")) {
            return "Needs Review";
        }

        if (containsAny(normalized,
                "pass", "passed", "looks good", "working", "works", "verified", "validated", "correct", "expected", "success", "successful")) {
            return "Pass / Verified";
        }

        return "Not Determined";
    }

    private static boolean containsAny(String normalizedText, String... phrases) {
        for (String phrase : phrases) {
            if (normalizedText.contains(normalizeTriggerText(phrase))) {
                return true;
            }
        }

        return false;
    }

    private static List<String> extractSummaryLines(List<String> cleanedLines) {
        List<String> selected = new ArrayList<>();
        Set<String> seen = new HashSet<>();

        for (String line : cleanedLines) {
            String normalized = normalizeTriggerText(line);

            if (normalized.isBlank() || seen.contains(normalized)) {
                continue;
            }

            if (isUsefulSummaryLine(normalized) || selected.size() < 2) {
                selected.add(shortenForSummary(line));
                seen.add(normalized);
            }

            if (selected.size() >= 4) {
                break;
            }
        }

        if (selected.isEmpty() && !cleanedLines.isEmpty()) {
            selected.add(shortenForSummary(cleanedLines.get(0)));
        }

        return selected;
    }

    private static boolean isUsefulSummaryLine(String normalizedText) {
        return containsAny(normalizedText,
                "validate", "verified", "confirm", "expected", "actual", "working", "not working", "error", "issue", "defect",
                "application", "salesforce", "banner", "submit", "status", "field", "document", "signature", "pdf", "test");
    }

    private static String shortenForSummary(String line) {
        String cleaned = line.replaceAll("\\s+", " ").trim();

        if (cleaned.length() <= 220) {
            return cleaned;
        }

        return cleaned.substring(0, 217).trim() + "...";
    }

    private static void addUserStoryHeading(XWPFDocument document, String headingText) {
        XWPFParagraph heading = document.createParagraph();
        XWPFRun headingRun = heading.createRun();

        headingRun.setBold(true);
        headingRun.setFontSize(18);
        headingRun.setText(headingText);
    }

    private static void addTestCaseHeading(XWPFDocument document, String headingText) {
        XWPFParagraph heading = document.createParagraph();
        XWPFRun headingRun = heading.createRun();

        headingRun.setBold(true);
        headingRun.setFontSize(14);
        headingRun.setText(headingText);
    }

    private static void addTranscriptReportSection(
            XWPFDocument document,
            List<TranscriptEntry> entries
    ) {
        XWPFParagraph pageBreak = document.createParagraph();
        pageBreak.setPageBreak(true);

        XWPFParagraph heading = document.createParagraph();
        XWPFRun headingRun = heading.createRun();

        headingRun.setBold(true);
        headingRun.setFontSize(16);
        headingRun.setText("Readable Teams Transcript Used for Screenshot Detection");


        for (TranscriptEntry entry : entries) {
            MatchResult matchResult = findMatchedTrigger(entry.text());

            addTranscriptSeparator(document);
        }
    }

    private static void addTranscriptSeparator(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(0);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Courier New");
        run.setFontSize(9);

    }

    static String formatSeconds(double seconds) {
        int totalMillis = (int) Math.round(seconds * 1000);

        int hours = totalMillis / 3_600_000;
        int minutes = (totalMillis % 3_600_000) / 60_000;
        int secs = (totalMillis % 60_000) / 1000;
        int millis = totalMillis % 1000;

        return String.format(
                Locale.ROOT,
                "%02d:%02d:%02d.%03d",
                hours,
                minutes,
                secs,
                millis
        );
    }

    private static int parseIntOrZero(String value) {
        if (value == null || value.isBlank()) {
            return 0;
        }

        return Integer.parseInt(value);
    }

    private static String removeBom(String text) {
        if (text == null) {
            return "";
        }

        return text.replace("\uFEFF", "");
    }

    private static String removeVttTags(String text) {
        if (text == null) {
            return "";
        }

        return text.replaceAll("<[^>]+>", "")
                .replace("&nbsp;", " ")
                .replace("&amp;", "&")
                .replace("&lt;", "<")
                .replace("&gt;", ">")
                .replaceAll("\\s+", " ")
                .trim();
    }

    private static void deleteRecursively(Path root) throws IOException {
        if (root == null || !Files.exists(root)) {
            return;
        }

        Files.walkFileTree(root, new SimpleFileVisitor<>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                    throws IOException {

                Files.deleteIfExists(file);
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult postVisitDirectory(Path dir, IOException exc)
                    throws IOException {

                Files.deleteIfExists(dir);
                return FileVisitResult.CONTINUE;
            }
        });
    }

    private static String resolveExecutable(String environmentVariable, String commandName) {
        String configuredPath = System.getenv(environmentVariable);

        if (configuredPath != null && !configuredPath.isBlank()) {
            return configuredPath;
        }

        String homebrewPrefix = "/opt/homebrew/bin/" + commandName;

        if (Files.isExecutable(Path.of(homebrewPrefix))) {
            return homebrewPrefix;
        }

        String intelHomebrewPrefix = "/usr/local/bin/" + commandName;

        if (Files.isExecutable(Path.of(intelHomebrewPrefix))) {
            return intelHomebrewPrefix;
        }

        return commandName;
    }

    private record ProcessingResult(
            List<TranscriptHit> hits,
            List<TranscriptEntry> entries,
            List<SectionTranscriptLine> sectionTranscriptLines
    ) {}

    private record TranscriptEntry(
            double seconds,
            String speaker,
            String text
    ) {}


    private record SectionTranscriptLine(
            String userStoryHeading,
            String testCaseHeading,
            double transcriptSeconds,
            String speaker,
            String text
    ) {}

    private record SummaryInfo(
            String status,
            List<String> summaryLines
    ) {}
    private record TranscriptHit(
            double transcriptSeconds,
            double captureSeconds,
            String speaker,
            String text,
            String matchedTrigger,
            String userStoryHeading,
            String testCaseHeading,
            Path imagePath
    ) {}

    private record MatchResult(
            boolean matched,
            String trigger
    ) {}

    private record TimestampInfo(
            double seconds,
            String speaker,
            String textBeforeTimestamp
    ) {}

    private record TestCaseMarker(int number, String heading) { }
}
