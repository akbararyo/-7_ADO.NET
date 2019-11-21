using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace LatihanADONET
{
    public partial class Form1 : Form
    {
        // constructor
        public Form1()
        {
            InitializeComponent();
            InisialisasiListView();
        }

        private void btnTesKoneksi_Click(object sender, EventArgs e)
        {
            //membuat objek connection
            OleDbConnection conn = GetOpenConnection();

            //cek status koneks
            if (conn.State == ConnectionState.Open) //koneksi berhasil
            {
                MessageBox.Show("Koneksi ke database berhasil !", "Informasi",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Koneksi ke database gagal !!!", "Informasi", 
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            conn.Dispose(); //tutup dan hapus objek connection dari memory
        }

        private OleDbConnection GetOpenConnection()
        {
            OleDbConnection conn = null; // deklarasi objek connection
            try // penggunaan blok try-catch untuk penanganan error
            {
                // atur ulang lokasi database yang disesuaikan dengan
                // lokasi database perpustakaan Anda
                string dbName = @"D:\18.11.2314\#7_ADO.NET\LatihanADO.NET\Database\DbPerpustakaan.mdb";

                //deklarasi variabel connectionString, ref: https://www.connection.com/
                string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", dbName);

                conn = new OleDbConnection(connectionString); //buat objek connection
                conn.Open(); //buka koneksi ke database
            }
            //jika terjadi error di blok try, akan tinggal ditangani langsung oleh blok catch
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error",
               MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
            return conn;
        }
        // atur format listview
        private void InisialisasiListView()
        {
            lvwMahasiswa.View = View.Details;
            lvwMahasiswa.FullRowSelect = true;
            lvwMahasiswa.GridLines = true;
            lvwMahasiswa.Columns.Add("No.", 30, HorizontalAlignment.Center);
            lvwMahasiswa.Columns.Add("NPM", 70, HorizontalAlignment.Center);
            lvwMahasiswa.Columns.Add("Nama", 190, HorizontalAlignment.Left);
            lvwMahasiswa.Columns.Add("Angkatan", 70, HorizontalAlignment.Center);
        }

        private void btnTampilkanData_Click(object sender, EventArgs e)
        {
            lvwMahasiswa.Items.Clear();
            // membuat objek Connection, sekaligus buka koneksi ke database
            OleDbConnection conn = GetOpenConnection();
            // deklarasi variabel sql untuk menampung perintah SELECT
            string sql = @"select npm, nama, angkatan from mahasiswa order by nama";
            // membuat objek Command untuk mengeksekusi perintah SQL
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            // membuat objek DataReader untuk menampung hasil perintah SELECT
            OleDbDataReader dtr = cmd.ExecuteReader(); // eksekusi perintah SELECT
            while (dtr.Read()) // gunakan perulangan utk menampilkan data kelistview
            {
                var noUrut = lvwMahasiswa.Items.Count + 1;
                var item = new ListViewItem(noUrut.ToString());
                item.SubItems.Add(dtr["npm"].ToString());
                item.SubItems.Add(dtr["nama"].ToString());
                item.SubItems.Add(dtr["angkatan"].ToString());
                lvwMahasiswa.Items.Add(item);
            }
            // setelah selesai digunakan,
            // segera hapus objek datareader, command dan connection dari memory
            dtr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            var result = 0;

            // validasi npm harus diisi
            if (txtNpmInsert.Text.Length == 0)
            {
                MessageBox.Show("NPM harus diisi !!!", "Informasi",
               MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
                txtNpmInsert.Focus();
                return;
            }

            // validasi nama harus diisi
            if (txtNamaInsert.Text.Length == 0)
            {
                MessageBox.Show("Nama harus diisi !!!", "Informasi",
               MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
                txtNamaInsert.Focus();
                return;
            }
            // membuat objek Connection, sekaligus buka koneksi ke database
            OleDbConnection conn = GetOpenConnection();
            // deklarasi variabel sql untuk menampung perintah INSERT
            var sql = @"insert into mahasiswa (npm, nama, angkatan)
 values (@npm, @nama, @angkatan)";
            // membuat objek Command untuk mengeksekusi perintah SQL
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            try
            {
                // set parameter untuk nama, angkatan dan npm
                cmd.Parameters.AddWithValue("@npm", txtNpmInsert.Text);
                cmd.Parameters.AddWithValue("@nama", txtNamaInsert.Text);
                cmd.Parameters.AddWithValue("@angkatan",
               txtAngkatanInsert.Text);
                result = cmd.ExecuteNonQuery(); // eksekusi perintah INSERT
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error",
               MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Dispose();
            }
            if (result > 0)
            {
                MessageBox.Show("Data mahasiswa berhasil disimpan !",
               "Informasi", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
                // reset form
                txtNpmInsert.Clear();
                txtNamaInsert.Clear();
                txtAngkatanInsert.Clear();
                txtNpmInsert.Focus();
            }
            else
                MessageBox.Show("Data mahasiswa gagal disimpan !!!",
               "Informasi", MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
            // setelah selesai digunakan,
            // segera hapus objek connection dari memory
            conn.Dispose();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            var result = 0;

            // validasi npm harus diisi
            if (txtNpmUpdate.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!", "Informasi", MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
                txtNpmUpdate.Focus();
                return;
            }

            // validasi nama harus diisi
            if (txtNamaUpdate.Text.Length == 0)
            {
                MessageBox.Show("Nama harus !!!", "Informasi", MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
                txtNamaUpdate.Focus();
                return;
            }
            // membuat objek Connection, sekaligus buka koneksi ke database
            OleDbConnection conn = GetOpenConnection();
            // deklarasi variabel sql untuk menampung perintah UPDATE
            string sql = @"update mahasiswa set nama = @nama, angkatan = @angkatan
 where npm = @npm";
            // membuat objek Command untuk mengeksekusi perintah SQL
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            try
            {
                // set parameter untuk nama, angkatan dan npm
                cmd.Parameters.AddWithValue("@nama", txtNamaUpdate.Text);
                cmd.Parameters.AddWithValue("@angkatan", txtAngkatanUpdate.Text);
                cmd.Parameters.AddWithValue("@npm", txtNpmUpdate.Text);
                result = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Dispose();
            }
            if (result > 0)
            {
                MessageBox.Show("Data mahasiswa berhasil diupdate !", "Informasi",
               MessageBoxButtons.OK,
                MessageBoxIcon.Information);
                // reset form
                txtNpmUpdate.Clear();
                txtNamaUpdate.Clear();
                txtAngkatanUpdate.Clear();
                txtNpmUpdate.Focus();
            }
            else
                MessageBox.Show("Data mahasiswa gagal diupdate !!!", "Informasi",
               MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
            // setelah selesai digunakan,
            // segera hapus objek connection dari memory
            conn.Dispose();
        }

        private void btnCariUpdate_Click(object sender, EventArgs e)
        {
            // validasi npm harus diisi
            if (txtNpmUpdate.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!", "Informasi", MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
                txtNpmUpdate.Focus();
                return;
            }
            // membuat objek Connection, sekaligus buka koneksi ke database
            OleDbConnection conn = GetOpenConnection();
            // deklarasi variabel sql untuk menampung perintah SELECT
            string sql = @"select npm, nama, angkatan
 from mahasiswa
where npm = @npm";
            // membuat objek Command untuk mengeksekusi perintah SQL
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            cmd.Parameters.AddWithValue("@npm", txtNpmUpdate.Text);
            // membuat objek DataReader untuk menampung hasil perintah SELECT
            OleDbDataReader dtr = cmd.ExecuteReader(); // eksekusi perintah SELECT
            if (dtr.Read()) // data ditemukan
            {
                // tampilkan nilainya ke textbox
                txtNpmUpdate.Text = dtr["npm"].ToString();
                txtNamaUpdate.Text = dtr["nama"].ToString();
                txtAngkatanUpdate.Text = dtr["angkatan"].ToString();
            }
            else
                MessageBox.Show("Data mahasiswa tidak ditemukan !", "Informasi",
               MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            // setelah selesai digunakan,
            // segera hapus objek datareader, command dan connection dari memory
            dtr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnCariDelete_Click(object sender, EventArgs e)
        {
            // validasi npm harus diisi
            if (txtNpmDelete.Text.Length == 0)
            {
                MessageBox.Show("NPM harus !!!", "Informasi", MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
                txtNpmDelete.Focus();
                return;
            }
            // membuat objek Connection, sekaligus buka koneksi ke database
            OleDbConnection conn = GetOpenConnection();
            // deklarasi variabel sql untuk menampung perintah SELECT
            string sql = @"select npm, nama, angkatan
 from mahasiswa
where npm = @npm";
            // membuat objek Command untuk mengeksekusi perintah SQL
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            cmd.Parameters.AddWithValue("@npm", txtNpmDelete.Text);
            // membuat objek DataReader untuk menampung hasil perintah SELECT
            OleDbDataReader dtr = cmd.ExecuteReader(); // eksekusi perintah SELECT
            if (dtr.Read()) // data ditemukan
            {
                // tampilkan nilainya ke textbox
                txtNpmDelete.Text = dtr["npm"].ToString();
                txtNamaDelete.Text = dtr["nama"].ToString();
                txtAngkatanDelete.Text = dtr["angkatan"].ToString();
            }
            else
                MessageBox.Show("Data mahasiswa tidak ditemukan !", "Informasi",
               MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            // setelah selesai digunakan,
            // segera hapus objek datareader, command dan connection dari memory
            dtr.Dispose();
            cmd.Dispose();
            conn.Dispose();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {

        }
    }
}
