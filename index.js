import { createReport } from 'docx-templates';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs/promises';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const templatePath = path.join(__dirname, 'TEMPLATEE.docx');
const outputPath = path.join('C:', 'Users', 'NazrinGuliyeva','Documents', 'output.docx');

async function main() {
  try {
    try {
      await fs.access(templatePath);
      console.log(`Template: ${templatePath}`);
    } catch {
      throw new Error(`Template file not found: ${templatePath}\nplease control file path`);
    }

    const reportBuffer = await createReport({
      template: await fs.readFile(templatePath),
      data: {
        meetingTitle: 'Test',
        venue:'Buta',
        recordedBy:'Sabina Asadova',
        type:'in person',
        date: new Date().toLocaleDateString('tr-TR'),
        agendaBy:'test',
        duration: '90 minute',
        time:'10:00PM',
        chairedBy:'test',
        attendees:"nazrin,sdhcbdhs,ahdcbhas,ahbchavch,jhdvcjad",
        apologies:"sdkhvgh,ksdgc,ksdcg,ksdhgbhs,skdjg"
        
      },
      cmdDelimiter: ['{{', '}}'] 
    });

    
    console.log(`writing file: ${outputPath}`);
    await fs.writeFile(outputPath, reportBuffer);
    
    
    try {
      await fs.access(outputPath);
      console.log(`succes create:\n${outputPath}`);
      console.log(`explorer.exe /select,"${outputPath}"`);
    } catch {
      throw new Error('write file but not create!');
    }

  } catch (error) {
    console.error('error:', error.message);

  }
}

main();