import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.*;
import javax.sound.sampled.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.filechooser.FileNameExtensionFilter;

public class LotteryApp extends JFrame {

    private final CardLayout cardLayout;
    private final JPanel mainPanel;
    private final JLabel productLabel;
    private final JLabel resultLabel;
    private final java.util.List<Item> items = new ArrayList<>();
    private int currentIndex = 0;
    private javax.swing.Timer rouletteTimer;
    private final Random random = new Random();
    private final java.util.List<Integer> availableNumbers = new ArrayList<>();

    private int drawDelay = 3000;
    private int volume = 50;

    private Point initialClick;

    private final JTextArea historyArea = new JTextArea();
    private Workbook historyWorkbook = new XSSFWorkbook();
    private Sheet historySheet = historyWorkbook.createSheet("抽選履歴");
    private int historyRowIndex = 0;

    // WAV用
    private Thread rollSoundThread;
    private Clip currentClip;

    public LotteryApp() {
        Row headerRow = historySheet.createRow(historyRowIndex++);
        headerRow.createCell(0).setCellValue("番号");
        headerRow.createCell(1).setCellValue("商品名");

        setUndecorated(true);
        GraphicsDevice gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
        gd.setFullScreenWindow(this);
        setDefaultCloseOperation(EXIT_ON_CLOSE);

        JPanel titleBar = new JPanel(new BorderLayout());
        titleBar.setBackground(new Color(60, 60, 60));
        titleBar.setPreferredSize(new Dimension(getWidth(), 40));

        JLabel titleLabel = new JLabel("抽選会", SwingConstants.CENTER);
        titleLabel.setForeground(Color.WHITE);
        titleLabel.setFont(new Font("Meiryo", Font.BOLD, 32));

        JButton closeButton = new JButton("✕");
        closeButton.setFocusPainted(false);
        closeButton.setBorderPainted(false);
        closeButton.setBackground(new Color(200, 50, 50));
        closeButton.setForeground(Color.WHITE);
        closeButton.setFont(new Font("Meiryo", Font.BOLD, 37));
        closeButton.addActionListener(e -> {
            saveHistoryExcel();
            System.exit(0);
        });

        titleBar.add(titleLabel, BorderLayout.CENTER);
        titleBar.add(closeButton, BorderLayout.EAST);

        titleBar.addMouseListener(new MouseAdapter() {
            public void mousePressed(MouseEvent e) {
                initialClick = e.getPoint();
            }
        });
        titleBar.addMouseMotionListener(new MouseMotionAdapter() {
            public void mouseDragged(MouseEvent e) {
                int thisX = getLocation().x;
                int thisY = getLocation().y;
                int xMoved = e.getX() - initialClick.x;
                int yMoved = e.getY() - initialClick.y;
                setLocation(thisX + xMoved, thisY + yMoved);
            }
        });

        cardLayout = new CardLayout();
        mainPanel = new JPanel(cardLayout);
        mainPanel.setBackground(new Color(230, 240, 255));

        JPanel loadPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);
        JLabel title = new JLabel("Excelファイルを選択してください（商品名・個数）");
        JButton loadButton = new JButton("ファイルを開く");
        loadButton.setPreferredSize(new Dimension(200, 60));
        loadButton.addActionListener(e -> openExcelChooser());
        gbc.gridy = 0;
        loadPanel.add(title, gbc);
        gbc.gridy = 1;
        loadPanel.add(loadButton, gbc);

        JPanel lotteryPanel = new JPanel(new BorderLayout(10, 10));
        lotteryPanel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));

        productLabel = new JLabel("次の賞：-", SwingConstants.CENTER);
        productLabel.setFont(new Font("Meiryo", Font.BOLD, 44));

        resultLabel = new JLabel("抽選結果はここに表示されます", SwingConstants.CENTER);
        resultLabel.setFont(new Font("Meiryo", Font.BOLD, 60));
        resultLabel.setForeground(new Color(30, 50, 120));

        JButton drawButton = new JButton("抽選");
        JButton exitButton = new JButton("終了");
        drawButton.setFont(new Font("Meiryo", Font.BOLD, 32));
        exitButton.setFont(new Font("Meiryo", Font.BOLD, 32));
        Dimension buttonSize = new Dimension(200, 80);
        drawButton.setPreferredSize(buttonSize);
        exitButton.setPreferredSize(buttonSize);
        drawButton.addActionListener(e -> startRoulette(drawButton));
        exitButton.addActionListener(e -> {
            saveHistoryExcel();
            System.exit(0);
        });

        JPanel buttonPanel = new JPanel();
        buttonPanel.add(drawButton);
        buttonPanel.add(exitButton);

        JPanel historyPanel = new JPanel(new BorderLayout());
        historyPanel.setPreferredSize(new Dimension(300, 0));
        historyArea.setEditable(false);
        historyArea.setFont(new Font("Meiryo", Font.PLAIN, 20));
        JScrollPane scrollPane = new JScrollPane(historyArea);
        historyPanel.add(new JLabel("抽選済みリスト", SwingConstants.CENTER), BorderLayout.NORTH);
        historyPanel.add(scrollPane, BorderLayout.CENTER);

        JPanel centerPanel = new JPanel(new BorderLayout());
        centerPanel.add(productLabel, BorderLayout.NORTH);
        centerPanel.add(resultLabel, BorderLayout.CENTER);
        centerPanel.add(buttonPanel, BorderLayout.SOUTH);

        lotteryPanel.add(centerPanel, BorderLayout.CENTER);
        lotteryPanel.add(historyPanel, BorderLayout.WEST);

        mainPanel.add(loadPanel, "load");
        mainPanel.add(lotteryPanel, "lottery");

        JPanel settingsPanel = new JPanel(new GridBagLayout());
        GridBagConstraints sgbc = new GridBagConstraints();
        sgbc.insets = new Insets(10, 10, 10, 10);
        sgbc.fill = GridBagConstraints.HORIZONTAL;

        JLabel timeLabel = new JLabel("抽選までの時間（秒）:");
        JSlider timeSlider = new JSlider(1, 10, drawDelay / 1000);
        timeSlider.setMajorTickSpacing(1);
        timeSlider.setPaintTicks(true);
        timeSlider.setPaintLabels(true);
        timeSlider.addChangeListener(e -> drawDelay = timeSlider.getValue() * 1000);

        JLabel volumeLabel = new JLabel("音量（0〜100）:");
        JSlider volumeSlider = new JSlider(0, 100, volume);
        volumeSlider.setMajorTickSpacing(20);
        volumeSlider.setPaintTicks(true);
        volumeSlider.setPaintLabels(true);
        volumeSlider.addChangeListener(e -> volume = volumeSlider.getValue());

        sgbc.gridy = 0; settingsPanel.add(timeLabel, sgbc);
        sgbc.gridy = 1; settingsPanel.add(timeSlider, sgbc);
        sgbc.gridy = 2; settingsPanel.add(volumeLabel, sgbc);
        sgbc.gridy = 3; settingsPanel.add(volumeSlider, sgbc);

        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setBackground(Color.GRAY);
        tabbedPane.addTab("抽選", mainPanel);
        tabbedPane.addTab("設定", settingsPanel);

        JPanel framePanel = new JPanel(new BorderLayout());
        framePanel.setBorder(BorderFactory.createLineBorder(new Color(80, 80, 80), 3));
        framePanel.add(titleBar, BorderLayout.NORTH);
        framePanel.add(tabbedPane, BorderLayout.CENTER);

        add(framePanel);
    }

    private void openExcelChooser() {
        GraphicsDevice gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
        gd.setFullScreenWindow(null);
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();
            loadExcelFromFile(file);
        }
        gd.setFullScreenWindow(this);
    }

    private void loadExcelFromFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            items.clear();
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                Cell nameCell = row.getCell(0);
                Cell countCell = row.getCell(1);
                if (nameCell == null || countCell == null) continue;
                String name = nameCell.getStringCellValue();
                int count = (int) countCell.getNumericCellValue();
                items.add(new Item(name, count));
            }

            if (!items.isEmpty()) {
                String rangeStr = JOptionPane.showInputDialog(this,
                        "抽選番号の範囲を入力してください（例: 1:100,200:300）", "1:100");
                availableNumbers.clear();
                if (rangeStr != null && !rangeStr.isBlank()) {
                    for (String range : rangeStr.split(",")) {
                        String[] parts = range.split(":");
                        if (parts.length == 2) {
                            int start = Integer.parseInt(parts[0].trim());
                            int end = Integer.parseInt(parts[1].trim());
                            for (int i = start; i <= end; i++) availableNumbers.add(i);
                        }
                    }
                }
                if (availableNumbers.isEmpty()) {
                    for (int i = 1; i <= 100; i++) availableNumbers.add(i);
                }

                updateProductLabel();
                cardLayout.show(mainPanel, "lottery");
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Excelの読み込みに失敗しました。");
        }
    }

    private void startRoulette(JButton drawButton) {
        if (currentIndex >= items.size()) {
            resultLabel.setText("すべての賞が終了しました！");
            return;
        }

        drawButton.setEnabled(false);
        resultLabel.setText("");

        //WAV再生に変更
        rollSoundThread = playLoopingSound("sounds/roll.wav");

        rouletteTimer = new javax.swing.Timer(100, e -> {
            int randomNum = availableNumbers.get(random.nextInt(availableNumbers.size()));
            resultLabel.setText("抽選中: " + randomNum);
        });
        rouletteTimer.start();

        javax.swing.Timer stopTimer = new javax.swing.Timer(drawDelay, e -> {
            rouletteTimer.stop();
            stopSound(rollSoundThread);
            playSound("sounds/rollend.wav");

            int finalNum = availableNumbers.remove(random.nextInt(availableNumbers.size()));
            Item currentItem = items.get(currentIndex);

            resultLabel.setText("<html><div style='text-align:center;'>"
                    + "当選番号: "
                    + "<span style='font-size:50px;'>" + finalNum + "</span><br>"
                    + "（" + currentItem.name + "）"
                    + "</div></html>");

            historyArea.append(finalNum + " : " + currentItem.name + "\n");

            Row row = historySheet.createRow(historyRowIndex++);
            row.createCell(0).setCellValue(finalNum);
            row.createCell(1).setCellValue(currentItem.name);

            drawButton.setEnabled(true);

            currentItem.count--;
            if (currentItem.count <= 0) {
                currentIndex++;
                if (currentIndex < items.size()) updateProductLabel();
                else productLabel.setText("すべての賞が終了しました！");
            } else {
                updateProductLabel();
            }
        });
        stopTimer.setRepeats(false);
        stopTimer.start();
    }

    private void updateProductLabel() {
        if (currentIndex < items.size()) {
            Item item = items.get(currentIndex);
            productLabel.setText("<html><div style='text-align:center;'>"
                    + "次の賞<br>"
                    + "<span style='font-size:40px;'>" + item.name + "</span><br>"
                    + "残り " + item.count + " 個"
                    + "</div></html>");
        }
    }

    private void saveHistoryExcel() {
        try (FileOutputStream fos = new FileOutputStream("抽選履歴.xlsx")) {
            historyWorkbook.write(fos);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "履歴の保存に失敗しました！");
        }
    }

    //WAVサウンド再生
    private Thread playLoopingSound(String path) {
        Thread thread = new Thread(() -> {
            try {
                InputStream is = getClass().getClassLoader().getResourceAsStream(path);
                if (is == null) {
                    System.err.println("サウンドが見つかりません: " + path);
                    return;
                }
                BufferedInputStream bis = new BufferedInputStream(is);
                AudioInputStream audioIn = AudioSystem.getAudioInputStream(bis);
                currentClip = AudioSystem.getClip();
                currentClip.open(audioIn);
                if (currentClip.isControlSupported(FloatControl.Type.MASTER_GAIN)) {
                    FloatControl volumeControl = (FloatControl) currentClip.getControl(FloatControl.Type.MASTER_GAIN);
                    float dB = (float) (Math.log(volume / 100.0) / Math.log(10.0) * 20.0);
                    volumeControl.setValue(dB);
                }

                // 🎵 無限ループ再生（途切れない）
                currentClip.loop(Clip.LOOP_CONTINUOUSLY);
                currentClip.start();

                // Clipを閉じるまでスレッドを維持
                while (!Thread.currentThread().isInterrupted() && currentClip.isRunning()) {
                    Thread.sleep(100);
                }

                audioIn.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
        thread.start();
        return thread;
    }


    private void stopSound(Thread soundThread) {
        try {
            if (currentClip != null && currentClip.isRunning()) {
                currentClip.stop();
                currentClip.close();
            }
        } catch (Exception ignored) {}
        if (soundThread != null && soundThread.isAlive()) {
            soundThread.interrupt();
        }
    }

    private void playSound(String path) {
        if (path.contains("rollend.wav")) stopSound(rollSoundThread);

        new Thread(() -> {
            try (InputStream is = getClass().getClassLoader().getResourceAsStream(path)) {
                if (is == null) {
                    System.err.println("サウンドファイルが見つかりません: " + path);
                    return;
                }
                BufferedInputStream bis = new BufferedInputStream(is);
                AudioInputStream audioIn = AudioSystem.getAudioInputStream(bis);
                Clip clip = AudioSystem.getClip();
                clip.open(audioIn);
                clip.start();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }).start();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new LotteryApp().setVisible(true));
    }
}

class Item {
    String name;
    int count;
    Item(String name, int count) {
        this.name = name;
        this.count = count;
    }
}


