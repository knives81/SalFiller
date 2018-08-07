package com.dashboard.salfiller;

import com.mashape.unirest.http.exceptions.UnirestException;
import com.sun.media.sound.InvalidFormatException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.IOException;

@SpringBootApplication
public class SalfillerApplication implements CommandLineRunner {

    String salPath = "C:\\Users\\pinzi\\myproject\\SAL.xlsm";
    String domain = "GICT_ITALY_MERCATO";
    String project = "P59_Change_Deploy_CRMT";
    String username = "mpinzi";
    String password = "gTlwd8";

	private static Logger LOG = LoggerFactory.getLogger(SalfillerApplication.class);

	@Autowired
	SalFillerService salFillerService;

	public static void main(String[] args) {
		LOG.info("STARTING THE APPLICATION");
		SpringApplication.run(SalfillerApplication.class, args);
	}

	@Override
	public void run(String... args) throws IOException, InvalidFormatException, org.apache.poi.openxml4j.exceptions.InvalidFormatException, UnirestException {
		LOG.info("EXECUTING.....");




        initParameter(args);

        salFillerService.getTestsetStatus(domain,project,username,password,salPath);
	}

    private void initParameter(String[] args) {

        if(args.length!=3 && args.length!=2) {
            System.out.println("Error input parameters!");
            System.out.println(help());
            System.exit(1);
        }
        username = args[0].split(":")[0];
        if(args[0].split(":").length==1) {
            password = "";
        } else {
            password = args[0].split(":")[1];
        }
        domain = args[1].split(":")[0];
        project = args[1].split(":")[1];

        if(args.length==2) {
            System.out.println("Working Directory = " + System.getProperty("user.dir"));
            salPath = System.getProperty("user.dir")+"\\SAL.xlsm";
        } else {
            salPath = args[2];
        }
        File f = new File(salPath);
        if(!f.exists()) {
            throw new RuntimeException("No SAL file found");
        }


    }

    private static String help() {
		return "use: java -jar salfiller.jar <username>:<password> <domain>:<project> <SAL path>\n"
				+ "<username> : Alm username\n"
				+ "<password> : Alm password\n"
				+ "<domain>   : Alm domain\n"
				+ "<project>  : Alm project\n"
				+ "<SAL path>   : SAL path - se mancante verrà usato SAL.xlsm nella directory corrente\n"
                + "example: java -jar salfiller.jar mpinzi:password123 GICT_ITALY_MERCATO:P32_CRM_T\"\n"
				+ "example: java -jar salfiller.jar mpinzi:password123 GICT_ITALY_MERCATO:P32_CRM_T \"C:\\users\\pinzi\\SAL.xlsm\"\n"
				+ "contacts: maurizio.pinzi@microfocus.com";
	}
}
