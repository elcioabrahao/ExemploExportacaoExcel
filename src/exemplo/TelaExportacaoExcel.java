package exemplo;

import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.Locale;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileSystemView;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class TelaExportacaoExcel extends JFrame {

	private JPanel contentPane;
	private JTextField textField;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					TelaExportacaoExcel frame = new TelaExportacaoExcel();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public TelaExportacaoExcel() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(null);
		setContentPane(contentPane);

		JLabel lblExportaoDeDados = new JLabel("Exportação de Dados de Para Arquivo Excel");
		lblExportaoDeDados.setBounds(66, 12, 362, 15);
		contentPane.add(lblExportaoDeDados);

		JButton btnExportar = new JButton("Exportar");
		btnExportar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// Ao clicar no botão chama método para exportar a planilha
				gerarArquivoExcel();

			}
		});
		btnExportar.setBounds(162, 57, 117, 25);
		contentPane.add(btnExportar);

		JLabel lblDestino = new JLabel("Destino:");
		lblDestino.setBounds(12, 151, 70, 15);
		contentPane.add(lblDestino);

		textField = new JTextField();
		textField.setBounds(83, 149, 331, 19);
		contentPane.add(textField);
		textField.setColumns(10);
	}

	private File getPastaParaSalvarArquivo() {

		// Exibe o file chooser e retorna um FILE correspondente a uma pasta de diretório
		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		jfc.setDialogTitle("Selecione a pasta para salvar seu arquivo: ");
		jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

		int returnValue = jfc.showSaveDialog(null);
		if (returnValue == JFileChooser.APPROVE_OPTION && jfc.getSelectedFile().isDirectory()) {

			return jfc.getSelectedFile();

		} else {
			return new File(".");
		}

	}

	private void gerarArquivoExcel() {

		// chama o File Chooser para poder escohar a pasta para guardar o arquivo
		File currDir = getPastaParaSalvarArquivo();
		String path = currDir.getAbsolutePath();
		// Adiciona ao nome da pasta o nome do arquivo que desejamos utilizar
		String fileLocation = path.substring(0, path.length()) + "/relatorio.xls";
		
		// mosta o caminho que exportamos na tela
		textField.setText(fileLocation);

		
		// Criação do arquivo excel
		try {
			
			// Diz pro excel que estamos usando portguês
			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("pt", "BR"));
			// Cria uma planilha
			WritableWorkbook workbook = Workbook.createWorkbook(new File(fileLocation), ws);
			// Cria uma pasta dentro da planilha
			WritableSheet sheet = workbook.createSheet("Pasta 1", 0);

			// Cria um cabeçario para a Planilha
			WritableCellFormat headerFormat = new WritableCellFormat();
			WritableFont font = new WritableFont(WritableFont.ARIAL, 16, WritableFont.BOLD);
			headerFormat.setFont(font);
			headerFormat.setBackground(Colour.LIGHT_BLUE);
			headerFormat.setWrap(true);

			Label headerLabel = new Label(0, 0, "Nome", headerFormat);
			sheet.setColumnView(0, 60);
			sheet.addCell(headerLabel);

			headerLabel = new Label(1, 0, "Idade", headerFormat);
			sheet.setColumnView(0, 40);
			sheet.addCell(headerLabel);

			// Cria as celulas com o conteudo
			WritableCellFormat cellFormat = new WritableCellFormat();
			cellFormat.setWrap(true);

			// Conteudo tipo texto
			Label cellLabel = new Label(0, 2, "Elcio Bonitão", cellFormat);
			sheet.addCell(cellLabel);
			// Conteudo tipo número (usar jxl.write... para não confundir com java.lang.number...)
			jxl.write.Number cellNumber = new jxl.write.Number(1, 2, 49, cellFormat);
			sheet.addCell(cellNumber);

			// Não esquecer de escrever e fechar a planilha
			workbook.write();
			workbook.close();

		} catch (IOException e1) {
			// Imprime erro se não conseguir achar o arquivo ou pasta para gravar
			e1.printStackTrace();
		} catch (WriteException e) {
			// exibe erro se acontecer algum tipo de celula de planilha inválida
			e.printStackTrace();
		}

	}
	

}
