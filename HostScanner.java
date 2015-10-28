import java.awt.*;
import java.awt.event.*;

import javax.swing.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Iterator;
import java.util.List;
import java.util.ArrayList;
import java.util.Arrays;

// XLSX Libraries
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings({"serial"})

public class HostScanner extends JPanel 
{
    private JButton btnScanHosts;
    private JButton btnBrowse;
    private JList<String> lstSearchResults;
    private JTextField tfSpreadsheet;
    private JLabel lblSpreadsheet;
    private JTextField tfSearch;
    private JLabel lblSearch;
    private JButton btnClear;

    private final JDialog dlg;

    private List<String> searchStrings = new ArrayList<String>();
    
    private String filePath;

    private void selectSpreadsheet()
    {
        FileDialog fd = new FileDialog((JFrame) null, "Select Host List", FileDialog.LOAD);
        fd.setDirectory(System.getProperty("user.dir"));

        if (System.getProperty("os.name").toLowerCase().indexOf("mac") >= 0)
        {
        	fd.setFilenameFilter(new FilenameFilter() 
        	{
        		@Override
        		public boolean accept(File dir, String name) 
        		{
        			return name.endsWith(".xlsx");
        		}
        	});
        }
        
        else
        {
        	fd.setFile("*.xlsx");
        }

        fd.setVisible(true);

        File fileWithPath = new File(fd.getDirectory() + fd.getFile());
        
        try
        {
        	String fullPath = fileWithPath.getCanonicalPath();
        	
        	if (fullPath != null && fullPath.length() != 0 && !fullPath.contains("null"))
        	{
        		// Save the spreadsheet path
        		this.filePath = fullPath;
        		tfSpreadsheet.setText(fileWithPath.getName());
        	}
        }
        
        catch (IOException ioe)
        {
        	System.out.println(ioe);	
        }
    }

    /* Method sends an HTTP GET request to a given host and returns the response as a String */
    private static String SendHTTPGetRequest(String hostAddress) throws Exception
    {
        final String USER_AGENT = "Mozilla/5.0";

		URL obj = new URL("http://" + hostAddress);
		HttpURLConnection con = (HttpURLConnection) obj.openConnection();
 
		// Optional default is GET
		con.setRequestMethod("GET");
 
		// Add request header
		con.setRequestProperty("User-Agent", USER_AGENT);
 
		BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
		String inputLine;
		StringBuffer response = new StringBuffer();
 
		while ((inputLine = in.readLine()) != null) 
        {
			response.append(inputLine);
		}

		in.close();

        return response.toString();
	}

    /* Method to search through a HTTP GET response for a specific string */
    private int searchForStrings(String HTTPResponse)
    {
        // NB: This currently returns true if the host matches ANY of the strings
        for (String stringToSearchFor : searchStrings)
        {
            if (HTTPResponse.toLowerCase().contains(stringToSearchFor.toLowerCase()))
            {
                /* Success, we found it! */
                return 0;
            }
        }

        return -1;
    }

    /* This method opens a CSV file containing the host IP addresses and reads them into memory. Throwing this in as a freebie
    private void getHostListCSV(String hostListPath, List<String> IPList) throws FileNotFoundException
    {
        Scanner CSVScanner = new Scanner(new File(hostListPath));
        CSVScanner.useDelimiter(",");

        while(CSVScanner.hasNext())
        {
            IPList.add(CSVScanner.next().trim());
        }

        CSVScanner.close();
    } */
    
    // This method opens an Excel XML (.xlsx) spreadsheet containing the host IP addresses and reads them into memory
    private void getHostList(String hostListPath, List<String> IPList) throws IOException
    {
    	FileInputStream excelFile = new FileInputStream(new File(hostListPath));
    	
    	try 
    	{
			XSSFWorkbook theWorkbook = new XSSFWorkbook(excelFile);
			
			XSSFSheet theSheet = theWorkbook.getSheetAt(0);
			
			Iterator<Row> rowIterator = theSheet.iterator();
			
			while (rowIterator.hasNext())
			{
				Row theRow = rowIterator.next();
				
				// For each row, iterate through the columns
				Iterator<Cell> cellIterator = theRow.cellIterator();
				
				while (cellIterator.hasNext())
				{
					Cell theCell = cellIterator.next();
					
					// Add the cell contents to the list of IP addresses
					IPList.add(theCell.getStringCellValue().trim());
				}
			}
			theWorkbook.close();
			excelFile.close();
		}
    	
    	catch (IOException ioe) 
    	{
    		JOptionPane.showMessageDialog(null, "The list of hosts could not be read!", "Error", JOptionPane.ERROR_MESSAGE);
    		System.out.println(ioe);
			return;
		}
    }

    // Main method to scan hosts 
    private void scanHosts(String hostListPath)
    {
        // Keep track of the number of hosts that could not be reached
        int hostsDown = 0;

        // Create list to store all addresses
        List<String> IPList = new ArrayList<String>();

        // Create list to store only matching addresses
        List<String> matchingIPList = new ArrayList<String>();

        // Import addresses from the file
        try
        {
            getHostList(hostListPath, IPList);
        }

        catch (Exception ex)
        {
            JOptionPane.showMessageDialog(null, "The list of hosts was not found or could not be read!", "Error", JOptionPane.ERROR_MESSAGE);
            System.out.println(ex);
            return;
        } 

        // For each IP address in the list of hosts
        for (String IP : IPList)
        {
            String HTTPResponse;

            // Send HTTP get request
            try
            {
                HTTPResponse = SendHTTPGetRequest(IP);
            }

            catch (Exception e)
            {
                //JOptionPane.showMessageDialog(null, "Error receiving HTTP GET response. Check your network connection", "Error", JOptionPane.ERROR_MESSAGE);
                //System.out.println(e);
                //return;
                hostsDown++;
                continue;
            }

            // Search for the search strings in the response
            if (searchForStrings(HTTPResponse) == 0)
            {
                // Add to the list of matching IP addresses if it matches
                matchingIPList.add(IP);
            }
        }

        // After loop, repopulate lstSearchResults with the contents of matchingIPList 
        DefaultListModel<String> model = new DefaultListModel<String>();

        if (hostsDown != 0)
        {
            model.addElement("Warning: " + hostsDown + " host(s) could not be reached.");
        }

        for (String IP : matchingIPList)
        {
            model.addElement(IP);
        }

        if (matchingIPList.size() == 0)
        {
            model.addElement("No matches found.");
        }

        lstSearchResults.setModel(model);
    }

    public HostScanner() 
    {
        // Initialise data
        String[] searchResults = {"Matching IP Addresses will appear here."};
        filePath = "";

        // Construct graphical components
        btnScanHosts  = new JButton ("Scan Hosts");
        btnBrowse = new JButton ("Browse...");
        lstSearchResults = new JList<String> (searchResults);
        tfSpreadsheet = new JTextField (5);
        tfSpreadsheet.setEditable(false);
        lblSpreadsheet = new JLabel ("Spreadsheet");
        tfSearch = new JTextField (5);
        lblSearch = new JLabel ("Search String(s)");
        btnClear = new JButton ("Clear");


        dlg = new JDialog((JFrame) null, "Scanning Hosts", true);
        JProgressBar dpb = new JProgressBar(0, 500);
        dpb.setIndeterminate(true);
        dlg.add(BorderLayout.CENTER, dpb);
        dlg.add(BorderLayout.NORTH, new JLabel("Scanning..."));
        dlg.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
        dlg.setSize(300, 75);
        dlg.setLocationRelativeTo(null);

        // Set component properties
        lstSearchResults.setToolTipText ("Search Results");
        tfSearch.setToolTipText("Multiple values can be entered, separated by commas");

        // Adjust component size and set layout
        setPreferredSize (new Dimension (500, 380));
        setLayout (null);

        // Add components
        add (btnScanHosts);
        add (btnBrowse);
        add (lstSearchResults);
        add (tfSpreadsheet);
        add (lblSpreadsheet);
        add (tfSearch);
        add (lblSearch);
        add (btnClear);

        // Set component bounds (only required when using Absolute Positioning)
        btnScanHosts.setBounds (10, 345, 480, 35);
        btnBrowse.setBounds (390, 20, 100, 20);
        lstSearchResults.setBounds (10, 85, 480, 250);
        tfSpreadsheet.setBounds (125, 20, 260, 20);
        lblSpreadsheet.setBounds (10, 15, 100, 25);
        tfSearch.setBounds (125, 45, 260, 20);
        lblSearch.setBounds (10, 40, 100, 25);
        btnClear.setBounds (390, 45, 100, 20);
        
        // Event Listeners
        btnClear.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent e)
            {
                tfSearch.setText("");
                lstSearchResults.setListData(new String[0]);
                tfSearch.requestFocus();
            }
        });

        btnBrowse.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent e)
            {
                selectSpreadsheet();
                tfSearch.requestFocus();
            }
        });

        btnScanHosts.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent e)
            {
                if (filePath.trim().equals(""))
                {
                    JOptionPane.showMessageDialog(null, "You need to specify the spreadsheet containing the IP addresses!", "Error", JOptionPane.ERROR_MESSAGE);
                    btnBrowse.requestFocus();
                    return;
                }

                if (tfSearch.getText().trim().equals(""))
                {
                    JOptionPane.showMessageDialog(null, "You did not enter a search string.", "Error", JOptionPane.ERROR_MESSAGE);
                    tfSearch.requestFocus();
                    return;
                }

                //dlg.setVisible(true);
                
                // Split and ignore whitespace 
                String search = tfSearch.getText();
                searchStrings = Arrays.asList(search.split("\\s*,\\s*"));

                // Scan the hosts
                scanHosts(filePath);
                
                tfSearch.requestFocus();
            }
        });
    }

    public static void main (String[] args) 
    {
        JFrame frame = new JFrame ("HostScanner by OMCS");
        frame.setDefaultCloseOperation (JFrame.EXIT_ON_CLOSE);
        frame.getContentPane().add (new HostScanner());
        frame.pack();
        frame.setVisible (true);
    }
}

