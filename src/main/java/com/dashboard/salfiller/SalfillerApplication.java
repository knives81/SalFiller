package com.dashboard.salfiller;

import com.mashape.unirest.http.exceptions.UnirestException;
import com.sun.media.sound.InvalidFormatException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class SalfillerApplication implements CommandLineRunner {

	private static Logger LOG = LoggerFactory.getLogger(SalfillerApplication.class);

	@Autowired
	SalFillerService salFillerService;

	public static void main(String[] args) {
		LOG.info("STARTING THE APPLICATION");
		SpringApplication.run(SalfillerApplication.class, args);
	}

	@Override
	public void run(String... args) throws IOException, InvalidFormatException, org.apache.poi.openxml4j.exceptions.InvalidFormatException, UnirestException {
		LOG.info("EXECUTING : command line runner");

		String salPath = "C:\\Users\\pinzi\\myproject\\SAL.xlsm";
		String domain = "GICT_ITALY_MERCATO";
		String project = "P59_Change_Deploy_CRMT";
		String username = "mpinzi";
		String password = "gTlwd8";

		salFillerService.getTestsetStatus(domain,project,username,password,salPath);

		if(args.length!=3) {
			System.out.println("Error input parameters!");
			System.out.println(help());
			System.exit(1);
		} else {
			username = args[0].split(":")[0];
			if(args[0].split(":").length==1) {
				password = "";
			} else {
				password = args[0].split(":")[1];
			}


			domain = args[1].split(":")[0];
			project = args[1].split(":")[1];

			salPath = args[2];
		}

		for (int i = 0; i < args.length; ++i) {
			LOG.info("args[{}]: {}", i, args[i]);
		}
	}

	private static String help() {
		return "use: java -jar salfiller.jar <username>:<password> <domain>:<project> <PMO> <S>\n"
				+ "<username> : Alm username\n"
				+ "<password> : Alm password\n"
				+ "<domain>   : Alm domain\n"
				+ "<project>  : Alm project\n"
				+ "<folder>   : SAL path\n"
				+ "example: java -jar defectclient.jar mpinzi:password123 GICT_ITALY_MERCATO:P32_CRM_T \"C:/users/pinzi/SAL.xlsm\"\n"
				+ "contacts: maurizio.pinzi@microfocus.com";
	}
}
