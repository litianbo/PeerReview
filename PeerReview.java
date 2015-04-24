import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class PeerReview {
	private static String USER_NAME = "csci201Spring2015"; // GMail user name
															// (just the
	// part before "@gmail.com")
	private static String PASSWORD = "usernamepassword"; // GMail password
	private static String RECIPIENT = "";

	public static void main(String[] args) {
		readInput();
		// sendEmail();

	}

	private static void readInput() {
		try {
			int[][] scores = new int[124][10];
			int[] numOfRe = new int[124];
			String[][] feedback = new String[124][10];
			int[] feedbackIndex = new int[124];
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					"201peerreview.xls"));

			// Get the workbook instance for XLS file
			HSSFWorkbook wb = new HSSFWorkbook(fs);

			// Get first sheet from the workbook
			HSSFSheet sheet = wb.getSheetAt(0);

			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();

			for (int rowNum = 0; rowNum < 118; rowNum++) {
				Row r = sheet.getRow(rowNum);
				int counter = 0;
				int userID = 0;
				// int lastColumn = Math.max(r.getLastCellNum(),
				// MY_MINIMUM_COLUMN_COUNT);
				int lastColumn = r.getLastCellNum();
				for (int cn = 0; cn < lastColumn; cn++) {
					Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);

					if (c == null) {
						// The spreadsheet is empty in this cell
						
						counter++;
						if (counter>26)
							break;
					} else {

						// Do something useful with the cell's contents
						if (c.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (counter == 0) {
								
								userID = (int) c.getNumericCellValue();
								numOfRe[userID]++;
								// System.out.println(userID);
							} else if (counter >= 11 && counter <= 18) {
								// System.out.print(counter + ", " + userID +
								// ".");
								if (userID == 122) {
									// System.out.println(c
									// .getNumericCellValue());
								}
								scores[userID][counter - 11] += (int) c
										.getNumericCellValue();
							} else if (counter == 25
									&& (int) c.getNumericCellValue() == 1) {
								counter = -1;
								userID = 0;
							} else if (counter == 25
									&& (int) c.getNumericCellValue() == 2) {
								userID = 0;
								break;
							}
							counter++;
						}
						if (c.getCellType() == Cell.CELL_TYPE_STRING) {
							if (counter == 19)
								feedback[userID][feedbackIndex[userID]++] = c
										.getStringCellValue();
							// System.out.print(userID);
							counter++;
						}

					}
				}
			}

			// System.out.println("");

			// FileOutputStream out = new FileOutputStream(new
			// File("output.xls"));
			for (int i = 122; i < 124; i++) {
				// System.out.print(i + "\t");
				for (int j = 0; j < 10; j++) {
					if (scores[i][j] != 0) {
						System.out.print("UserID: " + i + ", " + scores[i][j]
								+ ", NumberofReviews: " + numOfRe[i] + "\t");
					}
				}
				// System.out.println("");
			}
			System.out.println("");
			for (int i = 122; i < 124; i++) {
				// System.out.print(i + "\t");
				for (int j = 0; j < 10; j++) {
					if (feedback[i][j] != null) {
						System.out.print("UserID: " + i + ", " + feedback[i][j]
								+ ", NumberofReviews: " + numOfRe[i] + "\t");
					}
				}
				// System.out.println("");
			}
			HSSFWorkbook workbook2 = new HSSFWorkbook();
			HSSFSheet sheet2 = workbook2.createSheet("output");
			for (int i = 0; i < 124; i++) {
				Row row = sheet2.createRow(i);
				int j = 0;
				for (; j < 10; j++) {
					if (scores[i][j] != 0) {
						if (j < 7) {
							Cell cell = row.createCell(j);
							cell.setCellValue(((double) scores[i][j] / numOfRe[i]));
						} else {
							if (j == 7) {
								Cell cell = row.createCell(j);
								cell.setCellValue((double) scores[i][j]
										/ numOfRe[i]);
								Cell cell2 = row.createCell(j + 1);
								cell2.setCellValue(((double) scores[i][j] / numOfRe[i]) / 13 * 1.25);
								Cell cell3 = row.createCell(j + 2);
								cell3.setCellValue((((double) scores[i][j] / numOfRe[i]) / 13 * 1.25));
							}

						}
					}

				}
				for (int m = 0; m < feedbackIndex[i]; m++) {
					if (feedback[i][m] != null) {
						Cell cell = row.createCell(j + m);
						cell.setCellValue(feedback[i][m]);
					}
				}
			}
			try {
				FileOutputStream out = new FileOutputStream(new File(
						"output2.xls"));
				workbook2.write(out);
				out.close();
				System.out.println("Excel written successfully..");

			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			workbook2.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void sendEmail() {
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					"output.xls"));
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);
			HSSFRow row;
			HSSFCell cell;
			String[][] str = new String[120][22];
			int rows; // No of rows
			rows = sheet.getPhysicalNumberOfRows();

			int cols = 0; // No of columns
			int tmp = 0;

			// This trick ensures that we get the data properly even if it
			// doesn't start from first few rows
			for (int i = 0; i < 120 || i < rows; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					tmp = sheet.getRow(i).getPhysicalNumberOfCells();
					if (tmp > cols)
						cols = tmp;
				}
			}

			for (int r = 1; r < 121; r++) {
				row = sheet.getRow(r);
				if (row != null) {
					for (int c = 0; c < cols; c++) {
						cell = row.getCell((short) c);
						if (cell != null) {
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_NUMERIC:
								str[r - 1][c] = Double.toString(cell
										.getNumericCellValue());

								break;
							case Cell.CELL_TYPE_STRING:
								str[r - 1][c] = cell.getStringCellValue();

								break;
							}
						}
					}
					// System.out.println("");
				}

			}
			// System.out.print(str[80][2] + " ");
			String from = USER_NAME;
			String pass = PASSWORD;
			String[][] to = new String[120][2];

			for (int i = 0; i < 120; i++) {
				 to[i][0] = "csci201Spring2015@gmail.com";
				 to[i][1] = "csci201Spring2015@gmail.com";
				//to[i][0] = str[i][1];
				// to[i][0] = "csci477b@gmail.com";
				// to[i][1] = "tianboli@usc.edu";
				//to[i][1] = "jeffrey.miller@usc.edu";
			}
			String subject = "[CS201]Peer Review Feedback#1";
			String body[] = new String[120];
			for (int i = 0; i < 120; i++) {
				//
				//
				//
				//
				//
				//
				//
				//
				//
				//
				// Remember to change this
				body[i] = "Hi, "
						+ str[i][0]
						+ "\n"
						+ "Below is your feedback from peer review, "
						+ "\n"
						+ "Responsiveness to communication (out of 5) \n"
						+ str[i][3]
						+ "\n"
						+ "Willingness to work \n"
						+ str[i][4]
						+ "\n"
						+ "Amount of work completed \n"
						+ str[i][5]
						+ "\n"
						+ "Professionalism  \n"
						+ str[i][6]
						+ "\n"
						+ "On-time completion of assigned tasks	\n"
						+ str[i][7]
						+ "\n"
						+ "Quality of work completed \n"
						+ str[i][8]
						+ "\n"
						+ "Would you want to work with this person on a "
						+ "daily basis in a / professionial setting? \n"
						+ str[i][9]
						+ "\n"
						+ "If you had to assign a grade to this teammate based "
						+ "on the current / period of the project, what grade is it? (Out of 13) \n"
						+ str[i][10] + "\n" + "Overall Score (out of 2.5%) \n"
						+ str[i][12] + "\n";
				if (str[i][14] != null) {
					body[i] += "\n" + str[i][14] + "\n" + str[i][13];
				}
				if (str[i][15] != null) {
					body[i] += "\n" + str[i][15];
				}
				if (str[i][16] != null) {
					body[i] += "\n" + str[i][16];
				}
				if (str[i][17] != null) {
					body[i] += "\n" + str[i][17];
				}
				if (str[i][18] != null) {
					body[i] += "\n" + str[i][18];
				}
				if (str[i][19] != null) {
					body[i] += "\n" + str[i][19];
				}
				if (str[i][20] != null) {
					body[i] += "\n" + str[i][20];
				}
			}
			// System.out.println(to[0]);
			// comment out next line carefully
			sendFromGMail(from, pass, to, subject, body);
		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
		System.out.println("DONE");
	}

	private static void sendFromGMail(String from, String pass, String[][] to,
			String subject, String[] body1) {
		Properties props = System.getProperties();
		String host = "smtp.gmail.com";
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.host", host);
		props.put("mail.smtp.user", from);
		props.put("mail.smtp.password", pass);
		props.put("mail.smtp.port", "587");
		props.put("mail.smtp.auth", "true");
		String body = null;
		Session session = Session.getDefaultInstance(props);
		List<MimeMessage> messages = new ArrayList<MimeMessage>();
		for (int i = 0; i < 120; i++) {
			MimeMessage message = new MimeMessage(session);
			messages.add(message);
		}
		try {
			for (int i = 0; i < 120; i++)
				messages.get(i).setFrom(new InternetAddress(from));
			InternetAddress[] toAddress = new InternetAddress[2];
			Transport transport = session.getTransport("smtp");
			transport.connect(host, from, pass);
			// To get the array of addresses
			for (int i = 42; i < 43; i++) {
				toAddress[0] = new InternetAddress(to[i][0]);
				toAddress[1] = new InternetAddress(to[i][1]);
				body = body1[i];
				// forward to instructor
				messages.get(i).addRecipient(Message.RecipientType.TO,
						toAddress[0]);
				messages.get(i).addRecipient(Message.RecipientType.TO,
						toAddress[1]);
				messages.get(i).setSubject(subject);
				messages.get(i).setText(body);
				transport.sendMessage(messages.get(i), messages.get(i)
						.getAllRecipients());
			}
			transport.close();
			System.out.println("Finished");
		} catch (AddressException ae) {
			ae.printStackTrace();
		} catch (MessagingException me) {
			me.printStackTrace();
		}
	}
}
