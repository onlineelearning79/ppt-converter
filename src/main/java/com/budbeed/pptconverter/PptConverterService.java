package com.budbeed.pptconverter;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.springframework.stereotype.Service;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

@Service
public class PptConverterService
{
	private static final String UPLOADED_FOLDER = "./temp/";

	public String singleFileUpload(@RequestParam("file") MultipartFile file, RedirectAttributes redirectAttributes) throws Exception
	{

		// The path to the documents directory.

		if (file.isEmpty())
		{
			redirectAttributes.addFlashAttribute("message", "Please select a file to upload");
			return "redirect:uploadStatus";
		}

		try
		{
			// Get the file and save it somewhere
			Path path = Paths.get(UPLOADED_FOLDER + file.getOriginalFilename());
			Files.write(path, file.getBytes());
			convertToPngFiles(path.toFile(), 2);
			redirectAttributes.addFlashAttribute("message", "You successfully uploaded '" + file.getOriginalFilename() + "'");

		} catch (IOException e)
		{
			e.printStackTrace();
		}

		return "redirect:/uploadStatus";
	}

	public static void convertToPngFiles(File pptFile) throws Exception
	{
		convertToPngFiles(pptFile, 1);
	}

	public static void convertToPngFiles(File pptFile, float scale) throws Exception
	{

		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFile));
		Dimension pgsize = ppt.getPageSize();
		int width = (int) (pgsize.width * scale);
		int height = (int) (pgsize.height * scale);
		List<XSLFSlide> slide = ppt.getSlides();
		for (int i = 0; i < slide.size(); i++)
		{
			BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = img.createGraphics();

			// default rendering options
			graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
			graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
			graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
			graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

			graphics.setColor(Color.white);
			graphics.clearRect(0, 0, width, height);
			graphics.scale(scale, scale);

			// draw stuff
			slide.get(i).draw(graphics);

			// save the result
			File pngFile = new File(UPLOADED_FOLDER, i + ".png");
			System.out.println("Created image: " + pngFile.getName());
			FileOutputStream out = new FileOutputStream(pngFile);
			ImageIO.write(img, "png", out);
			//ppt.write(out);
			out.close();
		}
		System.out.println("Done");
	}

}
