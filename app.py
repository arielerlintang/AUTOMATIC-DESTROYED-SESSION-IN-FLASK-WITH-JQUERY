#perintah memanggil library flask
from flask import Flask, render_template, url_for, redirect, request, flash, session, get_flashed_messages #(url_for utk mewakilkan 127.5000)
from datetime import timedelta
import os
import hashlib
import mysql.connector


from werkzeug.utils import secure_filename   # <-- TAMBAHKAN
import openpyxl  
#konfgurasi awal flask
app=Flask(__name__)
app.secret_key = 'rahasia123'
#skrip konmeksi antara website dengan database
db=mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="trainit_bintang"
)
cursor = db.cursor(dictionary=True)
app.permanent_session_lifetime = timedelta(seconds=30)

#Rooting == untuk mengakses url dari web browser
@app.route('/')
def home():
    #skrip untuk ambil data dr database
    #cursor.execute("SELECT * FROM `dataset`")#string jadi pakai ""
    #pecah ke dalam bentuk array
    #dataset = cursor.fetchall()
    return render_template("home.html")

@app.route('/perhitungan')
def perhitungan():
    return render_template('perhitungan.html')

# Login Administrator
@app.route('/login', methods=['GET', 'POST'])
def login():
    # jika method adalah POST maka dapatkan inputan username dan password (request ke backend)
    if request.method == 'POST':
        username = request.form['username']
        password = hashlib.sha1(request.form['password'].encode()).hexdigest()

        # menampilkan data dari tabel admin berdasarkan username dan password
        cursor.execute("SELECT * FROM admin WHERE username=%s AND password=%s", (username, password))

        # pecah data dalam bentuk array
        admin = cursor.fetchone()

        # jika ada data admin maka simpan data admin ke dalam session
        if admin:
            session.permanent = True
            session['admin'] = admin
            flash('Berhasil Login', 'success')
            return redirect(url_for('home_admin'))
        
        # selain itu maka gagal login
        else:
            flash('Username atau Password salah', 'error')
            return redirect(url_for('login'))

    return render_template('login.html')

@app.route('/admin/home')
def home_admin():
    if 'admin' not in session:
        flash('Silakan login terlebih dahulu!', 'error')
        return redirect(url_for('login'))
    return render_template('admin/home.html', admin=session['admin'])


@app.route('/admin/dataset')
def admin_dataset():
    if 'admin' not in session:
        flash('Silakan login terlebih dahulu!', 'error')
        return redirect(url_for('login'))
    cursor.execute("SELECT * FROM dataset")
    dataset=cursor.fetchall()
    return render_template('admin/dataset.html', data=dataset)

@app.route('/admin/dataset/hapus/<int:id>')
def admin_dataset_hapus(id):
    #menghapus data dari tabel dataset berdasarkan id_dataset=2
    cursor.execute("DELETE FROM dataset where id_dataset=%s",(id,))
    db.commit()
    flash('Data berhasil dihapus','success')
    return redirect(url_for('admin_dataset'))


@app.route('/admin/dataset/ubah/<int:id>', methods=['GET','POST'])
def admin_dataset_ubah(id):
    #emenampilkan data dari tabel dataset berdasarkan id dataset yang id_dataset adalah 3
    cursor.execute("SELECT * FROM dataset where id_dataset=%s",(id,))
    dataset=cursor.fetchone()
    if request.method =='POST':
        #skrip utk mendapatkan inputan dari formulir == request.form['name']
        rank = request.form['rank']
        name = request.form['name']
        active = request.form['active']
        visits = request.form['visits']
        favourites = request.form['favourites']
        likes = request.form['likes']
        dislikes= request.form['dislikes']
        rating = request.form['rating']

        #simpan ke tabel dataset
        
        #jalankan perintah untuk mengeksekusi
        cursor.execute("UPDATE dataset SET rank=%s, name=%s, active=%s, visits=%s, favourites=%s, likes=%s, dislikes=%s, rating=%s WHERE id_dataset=%s", 
                       (rank, name, active, visits, favourites, likes, dislikes, rating, id))
        #untuk mengended
        db.commit()
        flash('Data berhasil diubah','success')
        #redirect ke halaman tampil
        return redirect(url_for('admin_dataset'))
    return render_template("admin/dataset_ubah.html",data=dataset)


#root tambah
@app.route('/admin/dataset_tambah', methods=['GET','POST']) #untuk menggunakan perintah get dan post
def admin_dataset_tambah(): #nama fungsi
    #jika ada request post maka ambil data dari formulir
    if request.method =='POST':
        #skrip utk mendapatkan inputan dari formulir == request.form['name']
        rank = request.form['rank']
        name = request.form['name']
        active = request.form['active']
        visits = request.form['visits']
        favourites = request.form['favourites']
        likes = request.form['likes']
        dislikes= request.form['dislikes']
        rating = request.form['rating']

        #simpan ke tabel dataset
        
        #jalankan perintah untuk mengeksekusi
        cursor.execute("INSERT INTO dataset (rank, name, active, visits, favourites, likes, dislikes, rating)" \
        "VALUES(%s, %s, %s, %s, %s, %s, %s, %s)",(rank,name,active,visits,favourites,likes,dislikes,rating))

        #untuk mengended
        db.commit()
        flash('Data berhasil ditambahkan','success')
        #redirect ke halaman tampil
        return redirect(url_for('admin_dataset'))

    return render_template ("admin/dataset_tambah.html") #nama halaman di templates

@app.route('/admin/dataset/import', methods=['POST'])
def import_dataset():
    file = request.files.get('file_xlsx')

    if file is None or file.filename == '':
        flash('File belum dipilih.', 'danger')
        return redirect(url_for('admin_dataset'))

    try:
        # langsung baca file upload-an tanpa simpan ke disk
        wb = openpyxl.load_workbook(file)
        ws = wb.active  # sheet pertama

        # baca header baris pertama
        header = [cell.value for cell in ws[1]]

        expected_header = [
            'Rank', 'Name', 'Active', 'Visits',
            'Favourites', 'Likes', 'Dislikes', 'Rating'
        ]

        # cek apakah heading sesuai
        if header != expected_header:
            flash('Format header file tidak sesuai template.', 'danger')
            return redirect(url_for('admin_dataset'))

        count = 0
        # iterasi mulai baris ke-2, hanya ambil value
        for row in ws.iter_rows(min_row=2, values_only=True):
            # kalau baris kosong, skip
            if all(v is None for v in row):
                continue

            rank, name, active, visits, favourites, likes, dislikes, rating = row

            # pakai cursor global yang sudah kamu definisikan di atas
            cursor.execute("""
                INSERT INTO dataset
                (rank, name, active, visits, favourites, likes, dislikes, rating)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                rank, name, active, visits,
                favourites, likes, dislikes, rating
            ))
            count += 1

        db.commit()
        flash(f'Berhasil mengimpor {count} baris data.', 'success')

    except Exception as e:
        db.rollback()
        flash(f'Terjadi kesalahan saat mengimpor data: {e}', 'danger')

    return redirect(url_for('admin_dataset'))

@app.route('/admin/profil/', methods=['GET','POST'])
def admin_profil():
    id_admin = session['admin']['id_admin']
    
    # ambil data dari tabel admin berdasarkan id_admin 
    cursor.execute("SELECT * FROM admin where id_admin=%s",(id_admin,))
    admin = cursor.fetchone()
    if request.method =='POST':
        #skrip utk mendapatkan inputan dari formulir == request.form['name']
        username = request.form['username']
        nama = request.form['nama']
        password = request.form['password']
        
        pe = hashlib.sha1(request.form['password'].encode()).hexdigest()

        
        # jika kosong password maka ubah data tanpa mengubah password
        if not password :
            cursor.execute("UPDATE admin SET username=%s, nama=%s, WHERE id_admin=%s", 
                       (username, nama, id_admin))
            
        # mengubah semua data dengan password
        else:
         cursor.execute("UPDATE admin SET username=%s, password=%s, nama=%s, WHERE id_admin=%s", 
                       (username, pe, nama, id_admin))
        
        
        #untuk mengended
        db.commit()
        flash('Data berhasil diubah','success')
        #redirect ke halaman tampil
        return redirect(url_for('admin_profil'))
    
    return render_template('admin/profil.html',data=admin)

@app.route('/admin/logout/')
def admin_logout():
    session.pop('admin', None)
    session.clear()
    return redirect(url_for('home'))



if __name__ == '__main__':
    #menjalankan server
    app.run(debug=True)
