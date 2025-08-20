#include <iostream>
#include <string>
#include <sstream>
#include <algorithm>
#include <cctype>
#include <limits>
#include <stdexcept>
#include <xlnt/xlnt.hpp>

using namespace std;

int trovaUltimaRigaConValore(const xlnt::worksheet &ws) {
    int ultimaRiga = 0;
    for (int riga = 1; riga <= ws.highest_row(); ++riga) {
        for (int colonna = 1; colonna <= ws.highest_column().index; ++colonna) {
            auto cella = ws.cell(xlnt::cell_reference(colonna, riga));
            if (cella.has_value()) {
                ultimaRiga = riga;
                break;
            }
        }
    }
    return ultimaRiga;
}

int trovaUltimaColonnaConValore(const xlnt::worksheet &ws) {
    int ultimaColonna = 0;
    for (int colonna = 1; colonna <= ws.highest_column().index; ++colonna) {
        for (int riga = 1; riga <= ws.highest_row(); ++riga) {
            auto cella = ws.cell(xlnt::cell_reference(colonna, riga));
            if (cella.has_value()) {
                ultimaColonna = colonna;
                break;
            }
        }
    }
    return ultimaColonna;
}

void stampaMatrice(const xlnt::worksheet &ws) {
    int ultimaRiga = trovaUltimaRigaConValore(ws);
    int ultimaColonna = trovaUltimaColonnaConValore(ws);
    cout << "Contenuto del foglio Excel:\n";
    for (int riga = 1; riga <= ultimaRiga; ++riga) {
        for (int colonna = 1; colonna <= ultimaColonna; ++colonna) {
            auto cella = ws.cell(xlnt::cell_reference(colonna, riga));
            cout << (cella.has_value() ? cella.to_string() : " ") << "\t";
        }
        cout << endl;
    }
}

static bool isInteger(const string &s) {
    if (s.empty()) return false;
    return all_of(s.begin(), s.end(), [](unsigned char ch){ return std::isdigit(ch) != 0; });
}

static string readQuotedOrToken(istringstream &iss) {
    while (iss.peek() == ' ') iss.get();
    if (iss.peek() == '"') {
        iss.get();
        string value;
        getline(iss, value, '"');
        return value;
    } else {
        string token;
        iss >> token;
        return token;
    }
}

static void copiaIntervalloDaFile(xlnt::worksheet &destWs,
        int destRiga, int destColonna,
        const string &srcPath, const string &sheetSpec,
        int srcRiga, int srcColonna,
        int numRighe, int numColonne) {

    xlnt::workbook srcWb;
    srcWb.load(srcPath);

    xlnt::worksheet srcWs;
    if (isInteger(sheetSpec)) {
        std::size_t idx = static_cast<std::size_t>(std::stoul(sheetSpec));
        if (idx == 0 || idx > srcWb.sheet_count()) {
            throw std::runtime_error("Indice foglio sorgente non valido");
        }
        srcWs = srcWb.sheet_by_index(idx - 1);
    } else {
        srcWs = srcWb.sheet_by_title(sheetSpec);
    }

    for (int r = 0; r < numRighe; ++r) {
        for (int c = 0; c < numColonne; ++c) {
            auto sCell = srcWs.cell(xlnt::cell_reference(srcColonna + c, srcRiga + r));
            auto dCell = destWs.cell(xlnt::cell_reference(destColonna + c, destRiga + r));
            if (sCell.has_value()) {
                dCell.value(sCell.value());
            } else {
                dCell.clear_value();
            }
        }
    }
}

void modificaCelleInterattiva(xlnt::workbook &wb, xlnt::worksheet &ws) {
    string input;

    cout << "Comandi: SET r c valore | COPIA_DA \"file.xlsx\" foglio r c righe colonne destR destC | ESPORTA \"file.xlsx\" | FOGLIO nome|indice | MOSTRA | AIUTO | Esci\n";

    while (true) {
        cout << "> ";
        getline(cin, input);

        if (input == "Esci") {
            break;
        }
        if (input == "" || input == "MOSTRA") {
            stampaMatrice(ws);
            continue;
        }
        if (input == "AIUTO") {
            cout << "Comandi:\n"
                 << "  SET r c valore\n"
                 << "  COPIA_DA \"file.xlsx\" foglio r c righe colonne destR destC\n"
                 << "  ESPORTA \"file.xlsx\"\n"
                 << "  FOGLIO nome|indice\n"
                 << "  MOSTRA\n"
                 << "  Esci\n";
            continue;
        }

        try {
            istringstream iss(input);
            string comando;
            iss >> comando;

            if (comando == "SET") {
                int riga, colonna;
                string nuovoValore;
                iss >> riga >> colonna;
                iss.ignore(numeric_limits<streamsize>::max(), ' ');
                getline(iss, nuovoValore);

                auto cella = ws.cell(xlnt::cell_reference(colonna, riga));
                cella.value(nuovoValore);
                cout << "OK" << endl;
            } else if (comando == "COPIA_DA") {
                string srcPath = readQuotedOrToken(iss);
                string sheetSpec = readQuotedOrToken(iss);
                int srcR, srcC, nR, nC, dstR, dstC;
                iss >> srcR >> srcC >> nR >> nC >> dstR >> dstC;
                copiaIntervalloDaFile(ws, dstR, dstC, srcPath, sheetSpec, srcR, srcC, nR, nC);
                cout << "Copiato." << endl;
            } else if (comando == "ESPORTA" || comando == "SALVA_COME") {
                string outPath = readQuotedOrToken(iss);
                wb.save(outPath);
                cout << "Salvato in: " << outPath << endl;
            } else if (comando == "FOGLIO") {
                string spec = readQuotedOrToken(iss);
                if (isInteger(spec)) {
                    std::size_t idx = static_cast<std::size_t>(std::stoul(spec));
                    if (idx == 0 || idx > wb.sheet_count()) {
                        throw std::runtime_error("Indice foglio non valido");
                    }
                    ws = wb.sheet_by_index(idx - 1);
                } else {
                    ws = wb.sheet_by_title(spec);
                }
                cout << "Foglio attivo: " << ws.title() << endl;
            } else {
                cerr << "Comando non riconosciuto. Digita AIUTO." << endl;
            }
        } catch (const std::exception &e) {
            cerr << "Errore: " << e.what() << endl;
        } catch (...) {
            cerr << "Errore sconosciuto." << endl;
        }
    }
}

int main() {
    string filePath;

    cout << "Inserisci il percorso del file Excel: ";
    getline(cin, filePath);

    xlnt::workbook wb;
    try {
        wb.load(filePath);
    } catch (const xlnt::exception &e) {
        cerr << "Errore durante il caricamento del file: " << e.what() << endl;
        return 1;
    }

    auto ws = wb.active_sheet();

    cout << "Matrice iniziale:\n";
    stampaMatrice(ws);

    modificaCelleInterattiva(wb, ws);

    try {
        wb.save(filePath);
        cout << "File Excel salvato con successo!" << endl;
    } catch (const xlnt::exception &e) {
        cerr << "Errore durante il salvataggio del file: " << e.what() << endl;
        return 1;
    }

    return 0;
}
