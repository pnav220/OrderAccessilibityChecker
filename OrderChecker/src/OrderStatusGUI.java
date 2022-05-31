import java.awt.EventQueue;
import javax.swing.*;
import java.awt.event.*;
import javax.swing.filechooser.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.awt.Color;

public class OrderStatusGUI implements ActionListener{

	private JFrame frame;
	private static JLabel label;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					OrderStatusGUI window = new OrderStatusGUI();
					window.frame.setVisible(true);
					JButton button = new JButton("Open");
					button.addActionListener(window);
			        JPanel panel = new JPanel();
			        panel.setBackground(Color.WHITE);
			        panel.add(button);
			        label = new JLabel("File Not Selected");
			        label.setForeground(Color.BLACK);
			        panel.add(label);
			        window.frame.getContentPane().add(panel);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public OrderStatusGUI() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame("Order Status Checker");
		frame.setBounds(300, 300, 300, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		JFileChooser chooser = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		chooser.setAcceptAllFileFilterUsed(false);
		chooser.setDialogTitle("Select a .xlsx file");
		FileNameExtensionFilter filter = new FileNameExtensionFilter(".xlsx files only!","xlsx");
		chooser.addChoosableFileFilter(filter);
		int i = chooser.showOpenDialog(null);
		if(i == JFileChooser.APPROVE_OPTION) {
			label.setText(chooser.getSelectedFile().getAbsolutePath());
            ExcelToText text = new ExcelToText(chooser.getSelectedFile());
            text.dumpCodes(chooser.getSelectedFile());
            try {
				text.checkStatus();
			} catch (EncryptedDocumentException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (InvalidFormatException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
            label.setText("Material Status Checked!");
		}
		else
			label.setText("Operation has been cancelled.");
	}

}
