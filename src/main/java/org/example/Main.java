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

public class Main {

    private static final double CAPTURE_OFFSET_SECONDS = 0.0;
    private static final double MIN_HIT_GAP_SECONDS = 0.5;

    private static final String FFMPEG_ENV = "FFMPEG_PATH";

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
                addTitle(document, "Teams Recording Screenshots");

                addSummary(document, result);

                if (result.hits().isEmpty()) {
                    addParagraph(document, "No screenshot triggers found in the Teams transcript.");
                } else {
                    for (int index = 0; index < result.hits().size(); index++) {
                        TranscriptHit hit = result.hits().get(index);
                        addScreenshotEntry(document, index + 1, hit, hit.imagePath());
                    }
                }

                addTranscriptReportSection(document, result.entries());

                try (OutputStream outputStream = Files.newOutputStream(outputDocx)) {
                    document.write(outputStream);
                }
            }

            System.out.println("Saved Word document to: " + outputDocx);

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

        System.out.println("Total transcript entries found: " + entries.size());

        for (TranscriptEntry entry : entries) {
            MatchResult matchResult = findMatchedTrigger(entry.text());

            System.out.println("Time      : " + formatSeconds(entry.seconds()));
            System.out.println("Speaker   : " + entry.speaker());
            System.out.println("Text      : " + entry.text());
            System.out.println("Normalized: " + normalizeTriggerText(entry.text()));
            System.out.println("Match     : " + (matchResult.matched() ? "YES - " + matchResult.trigger() : "NO"));
            System.out.println("--------------------------------------------------");

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
                            screenshotPath
                    ));
                }
            }
        }

        return new ProcessingResult(hits, entries);
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

        Path debugFile = createDebugTranscriptPath(outputDocx);
        Files.writeString(debugFile, content, StandardCharsets.UTF_8);
        System.out.println("Extracted transcript text saved to: " + debugFile);

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

    private static void addSummary(XWPFDocument document, ProcessingResult result) {
        addParagraph(document, "Total transcript entries found: " + result.entries().size());
        addParagraph(document, "Total screenshots captured: " + result.hits().size());
        addParagraph(document, "Capture offset seconds: " + CAPTURE_OFFSET_SECONDS);
        addParagraph(document, "Minimum hit gap seconds: " + MIN_HIT_GAP_SECONDS);
    }

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
        headingRun.setText("Screenshot " + index);

        addParagraph(document, "Transcript Time: " + formatSeconds(hit.transcriptSeconds()));
        addParagraph(document, "Captured Video Time: " + formatSeconds(hit.captureSeconds()));
        addParagraph(document, "Speaker: " + hit.speaker());
        addParagraph(document, "Matched Trigger: " + hit.matchedTrigger());
        addParagraph(document, "Transcript Text: " + hit.text());

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

        addParagraph(document, "This section shows each transcript entry, normalized text, and whether it matched a screenshot trigger.");

        for (TranscriptEntry entry : entries) {
            MatchResult matchResult = findMatchedTrigger(entry.text());

            addTranscriptSeparator(document);
            addTranscriptLine(document, "Time: " + formatSeconds(entry.seconds()));
            addTranscriptLine(document, "Speaker: " + entry.speaker());
            addTranscriptLine(document, "Original Transcript Text: " + entry.text());
            addTranscriptLine(document, "Normalized Text: " + normalizeTriggerText(entry.text()));
            addTranscriptLine(document, "Matched Trigger: " + (
                    matchResult.matched()
                            ? "YES - " + matchResult.trigger()
                            : "NO"
            ));
        }
    }

    private static void addTranscriptSeparator(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(0);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Courier New");
        run.setFontSize(9);
        run.setText("============================================================");
    }

    private static void addTranscriptLine(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(0);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Courier New");
        run.setFontSize(9);
        run.setText(text == null ? "" : text);
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
            List<TranscriptEntry> entries
    ) {}

    private record TranscriptEntry(
            double seconds,
            String speaker,
            String text
    ) {}

    private record TranscriptHit(
            double transcriptSeconds,
            double captureSeconds,
            String speaker,
            String text,
            String matchedTrigger,
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
}