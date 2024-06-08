using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using HtmlAgilityPack;
using OfficeOpenXml;

namespace PROJETO_2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // Defina o contexto da licença no construtor ou em outro ponto inicial do aplicativo
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Inicialização ao carregar o formulário, se necessário
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Arquivo de Texto (*.txt)|*.txt";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                string htmlContent = File.ReadAllText(filePath);

                List<ClientInfo> clients = ExtractClientInfo(htmlContent);
                listBox1.Items.Clear();
                foreach (var client in clients)
                {
                    listBox1.Items.Add($"{client.Name} - {client.Phone} - {client.Technician}");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Arquivo Excel (*.xlsx)|*.xlsx";
            saveFileDialog.FileName = "clientes.xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                List<ClientInfo> clients = listBox1.Items.Cast<string>()
                    .Select(item =>
                    {
                        var parts = item.Split(new[] { " - " }, StringSplitOptions.None);
                        return new ClientInfo
                        {
                            Name = parts[0],
                            Phone = parts[1],
                            Technician = parts.Length > 2 ? parts[2] : "Sem Técnico ainda"
                        };
                    })
                    .ToList();

                SaveToExcel(clients, filePath);
                MessageBox.Show("Lista salva com sucesso!");
            }
        }

        private List<ClientInfo> ExtractClientInfo(string htmlContent)
        {
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlContent);

            var clients = new List<ClientInfo>();

            var clientNodes = doc.DocumentNode.SelectNodes("//td[@abbr='cliente.razao']/div[contains(@style, 'text-align: left; width: 198px;')]");
            var phoneNodes = doc.DocumentNode.SelectNodes("//td[@abbr='cliente.telefone_celular']/div[contains(@style, 'text-align: left; width: 80px;')]");
            var technicianNodes = doc.DocumentNode.SelectNodes("//td[@abbr='funcionarios.funcionario']//div[@style='text-align: left; width: 76px;']");

            if (clientNodes != null && phoneNodes != null && technicianNodes != null)
            {
                int count = Math.Min(clientNodes.Count, Math.Min(phoneNodes.Count, technicianNodes.Count));

                for (int i = 0; i < count; i++)
                {
                    string clientName = clientNodes[i].InnerText.Trim();
                    string phone = phoneNodes[i].InnerText.Trim();
                    string technician = "Sem Técnico ainda";

                    if (technicianNodes.Count > i)
                    {
                        var technicianNode = technicianNodes[i];
                        if (!technicianNode.InnerText.Trim().Equals("&nbsp;", StringComparison.InvariantCultureIgnoreCase))
                        {
                            technician = technicianNode.Attributes["title"]?.Value.Trim() ?? "Sem Técnico ainda";
                        }
                    }

                    clients.Add(new ClientInfo { Name = clientName, Phone = phone, Technician = technician });
                }
            }

            return clients;
        }

        private void SaveToExcel(List<ClientInfo> clients, string filePath)
        {
            // Defina o contexto da licença
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Clientes");
                worksheet.Cells[1, 1].Value = "Nome";
                worksheet.Cells[1, 2].Value = "Telefone";
                worksheet.Cells[1, 3].Value = "Técnico";

                for (int i = 0; i < clients.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = clients[i].Name;
                    worksheet.Cells[i + 2, 2].Value = clients[i].Phone;
                    worksheet.Cells[i + 2, 3].Value = clients[i].Technician;
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }

    public class ClientInfo
    {
        public string Name { get; set; }
        public string Phone { get; set; }
        public string Technician { get; set; }
    }
}
