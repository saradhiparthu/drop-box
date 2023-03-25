package com.dropbox;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import com.dropbox.core.DbxApiException;
import com.dropbox.core.DbxException;
import com.dropbox.core.DbxRequestConfig;
import com.dropbox.core.v2.DbxClientV2;
import com.dropbox.core.v2.files.FileMetadata;
import com.dropbox.core.v2.users.FullAccount;

@SpringBootApplication
@RestController
public class DropBoxApplication {

	public static void main(String[] args) {
		SpringApplication.run(DropBoxApplication.class, args);

	}

	@RequestMapping(method = RequestMethod.POST, path = "/saveExcelInDropBox")
	public ResponseEntity<?> saveExcelInDropBox(@RequestBody GithubUser[] users,
			@RequestHeader(value = "ACCESS_TOKEN", required = true) String token) {
		try {
			// Drop-Box integration
			DbxRequestConfig config = DbxRequestConfig.newBuilder("dropbox/java-tutorial").build();
			DbxClientV2 client = new DbxClientV2(config, token);
			// Get current account info
			FullAccount account = client.users().getCurrentAccount();
			System.out.println(account.getName().getDisplayName());
			// Blank workbook using Apache POI
			// create Workbook
			XSSFWorkbook workbook = new XSSFWorkbook();

			// Create a blank sheet
			XSSFSheet sheet = workbook.createSheet("GithubUser Data");
			int rownum = 0;
			System.out.println("Total Users:: " + users.length);
			for (int i = 0; i < users.length; i++) {
				// create Row
				Row row = sheet.createRow(rownum++);

				Cell idCell = row.createCell(1);
				idCell.setCellValue(users[i].getId());

				Cell loginCell = row.createCell(2);
				loginCell.setCellValue(users[i].getLogin());
			}

			ByteArrayOutputStream bos = new ByteArrayOutputStream();
			try {
				workbook.write(bos);
				workbook.close();
			} finally {
				bos.close();
			}
			// Apache POI and Drop-Box
			byte[] bytes = bos.toByteArray();
			try (InputStream in = new ByteArrayInputStream(bytes)) {
				FileMetadata metadata = client.files().uploadBuilder("/GithubUsers.xlsx").uploadAndFinish(in);
				System.out.println(metadata.getSize());
			}
			System.out.println("GithubUsers.xlsx written successfully on disk.");
		} catch (DbxApiException e) {
			e.printStackTrace();
		} catch (DbxException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return ResponseEntity.status(HttpStatus.OK).body(users);
	}

}

class GithubUser {

	private String login;
	private Integer id;

	public String getLogin() {
		return login;
	}

	public void setLogin(String login) {
		this.login = login;
	}

	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

}