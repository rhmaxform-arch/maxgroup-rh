from app import app, db, Empresa, EMPRESAS_INICIAIS, migrate_db

with app.app_context():
    db.create_all()
    migrate_db()
    if Empresa.query.count() == 0:
        for nome in EMPRESAS_INICIAIS:
            db.session.add(Empresa(nome=nome))
        db.session.commit()

if __name__ == '__main__':
    app.run()
