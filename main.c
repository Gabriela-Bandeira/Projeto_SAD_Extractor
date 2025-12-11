#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <dirent.h>
#include <sys/stat.h>

#ifdef _WIN32
#include <windows.h>
#include <direct.h>
#define mkdir _mkdir
#define strncasecmp _strnicmp
#else
#include <unistd.h>
#endif

#include <tesseract/capi.h>
#include <leptonica/allheaders.h>
#include <xlsxwriter.h>


// -------------------------------------------------------------
// Função portátil: busca substring ignorando maiúsc/min
// -------------------------------------------------------------
char *stristr(const char *haystack, const char *needle) {

    if (!haystack || !needle) return NULL;

    size_t nlen = strlen(needle);

    for (; *haystack; haystack++) {
        if (strncasecmp(haystack, needle, nlen) == 0)
            return (char *)haystack;
    }
    return NULL;
}


// -------------------------------------------------------------
// Extrai campo após palavra-chave
// -------------------------------------------------------------
void extrair_campo(const char *texto, const char *chave, char *saida) {

    char *p = stristr(texto, chave);
    if (!p) { strcpy(saida, "N/A"); return; }

    p += strlen(chave);

    while (*p == ':' || *p == ' ') p++;

    int i = 0;
    while (*p && *p != '\n' && i < 255)
        saida[i++] = *p++;

    saida[i] = '\0';
}


// -------------------------------------------------------------
// OCR via Tesseract + Leptonica
// -------------------------------------------------------------
char *ocr_png(const char *png_path) {

    PIX *img = pixRead(png_path);
    if (!img) {
        fprintf(stderr, "Erro lendo PNG '%s'\n", png_path);
        return NULL;
    }

    TessBaseAPI *api = TessBaseAPICreate();
    TessBaseAPIInit3(api, NULL, "por");

    TessBaseAPISetImage2(api, img);

    char *out = TessBaseAPIGetUTF8Text(api);

    // --- salvar OCR em arquivo txt para depuração ---
    char txt_name[256];
    snprintf(txt_name, sizeof(txt_name), "%s.txt", png_path);

    FILE *ft = fopen(txt_name, "w");
    if (ft) {
        fputs(out, ft);
        fclose(ft);
    } else
        printf("Não consegui criar arquivo %s\n", txt_name);

    TessBaseAPIDelete(api);
    pixDestroy(&img);

    return out;
}


// -------------------------------------------------------------
// Converte PDF → PNG via pdftoppm
// -------------------------------------------------------------
char *extrair_texto_pdf(const char *arquivo_pdf) {

    static char comando[512];
    static char foto[256];

    char *texto_total = malloc(2000);
    texto_total[0] = '\0';
    size_t tam = 2000;

    int pagina = 1;

    while (1) {

        snprintf(comando, sizeof(comando),
                 "pdftoppm \"%s\" tmp_pag -png -f %d -l %d >nul 2>&1",
                 arquivo_pdf, pagina, pagina);

        system(comando);

        // <<< AQUI ESTAVA O ERRO: faltava %02d >>>
        snprintf(foto, sizeof(foto), "tmp_pag-%02d.png", pagina);

        FILE *t = fopen(foto, "rb");
        if (!t) break;  // acabou as páginas
        fclose(t);

        char *texto_pag = ocr_png(foto);
        if (texto_pag) {

            tam += strlen(texto_pag) + 1000;
            texto_total = realloc(texto_total, tam);
            strcat(texto_total, texto_pag);

            free(texto_pag);
        }

        remove(foto); // remove a imagem depois

        pagina++;
    }

    return texto_total;
}


// -------------------------------------------------------------
// Estrutura com campos extraídos
// -------------------------------------------------------------
typedef struct {
    char laudo[256];
    char endereco[256];
    char numero[64];
    //char complemento[?];
    char bairro[128];
    char cep[64];
    //char regiao[?];
    char cidade[128];
    char estado[64];
    char pais[64];
    //char confrontante_frente[?];
    //char confrontante_fundo[?];
    //char confrontante_Lesq[?];
    //char confrontante_Ldir[?];
    //char ponto_referencia[?];
    //char coord_geo_S[?];
    //char coor_geo_W[?];
    //char observacao[?];
    char area_terreno[64];
    char area_construida[64];
    char unidade_medida[64];
    char estado_conservacao[128];
    //char limitacao_administrativa[?];
    char criterio_valoracao[128];
    char data_valoracao[64];
    //char no_documento[?];
    //char valor_construcao_nova[?];
    //char valor_area_construida[?];
    char valor_total[64];
} DadosExtraidos;


// -------------------------------------------------------------
// Processa 1 PDF
// -------------------------------------------------------------
DadosExtraidos processar_laudo(const char *caminho) {

    DadosExtraidos d = {0};
    strcpy(d.laudo, caminho);

    char *texto = extrair_texto_pdf(caminho);
    if (!texto) return d;

    extrair_campo(texto, "Endereço", d.endereco);
    extrair_campo(texto, "Nº", d.numero);
    extrair_campo(texto, "Bairro", d.bairro);
    extrair_campo(texto, "CEP", d.cep);
    extrair_campo(texto, "Cidade", d.cidade);
    extrair_campo(texto, "Estado", d.estado);
    extrair_campo(texto, "País", d.pais);

    extrair_campo(texto, "Área do terreno", d.area_terreno);
    extrair_campo(texto, "Área construída", d.area_construida);
    extrair_campo(texto, "Unidade de medida", d.unidade_medida);
    extrair_campo(texto, "Estado de conservação", d.estado_conservacao);

    extrair_campo(texto, "Critério de valoração", d.criterio_valoracao);
    extrair_campo(texto, "Data da valoração", d.data_valoracao);
    extrair_campo(texto, "Valor total", d.valor_total);

    free(texto);
    return d;
}


// -------------------------------------------------------------
// Gera planilha XLSX
// -------------------------------------------------------------
void gerar_planilha(DadosExtraidos *dados, int n) {

    mkdir("output");

    lxw_workbook *wb = workbook_new("output/resultado.xlsx");
    lxw_worksheet *ws = workbook_add_worksheet(wb, NULL);

    const char *cab[] = {
        "Laudo","Endereço","Número","Bairro","CEP","Cidade","Estado",
        "País","Área Terreno","Área Construída","Unidade",
        "Conservação","Critério","Data Valoração","Valor Total"
    };

    for (int i = 0; i < 15; i++)
        worksheet_write_string(ws, 0, i, cab[i], NULL);

    for (int i = 0; i < n; i++) {
        int r = i + 1;

        worksheet_write_string(ws, r, 0, dados[i].laudo, NULL);
        worksheet_write_string(ws, r, 1, dados[i].endereco, NULL);
        worksheet_write_string(ws, r, 2, dados[i].numero, NULL);
        worksheet_write_string(ws, r, 3, dados[i].bairro, NULL);
        worksheet_write_string(ws, r, 4, dados[i].cep, NULL);
        worksheet_write_string(ws, r, 5, dados[i].cidade, NULL);
        worksheet_write_string(ws, r, 6, dados[i].estado, NULL);
        worksheet_write_string(ws, r, 7, dados[i].pais, NULL);
        worksheet_write_string(ws, r, 8, dados[i].area_terreno, NULL);
        worksheet_write_string(ws, r, 9, dados[i].area_construida, NULL);
        worksheet_write_string(ws, r, 10, dados[i].unidade_medida, NULL);
        worksheet_write_string(ws, r, 11, dados[i].estado_conservacao, NULL);
        worksheet_write_string(ws, r, 12, dados[i].criterio_valoracao, NULL);
        worksheet_write_string(ws, r, 13, dados[i].data_valoracao, NULL);
        worksheet_write_string(ws, r, 14, dados[i].valor_total, NULL);
    }

    workbook_close(wb);
    printf("Planilha salva em output/resultado.xlsx\n");
}


// -------------------------------------------------------------
// MAIN
// -------------------------------------------------------------
int main() {

    DIR *d = opendir("data");
    if (!d) {
        printf("Crie uma pasta chamada 'data' e coloque os PDFs dentro.\n");
        return 1;
    }

    struct dirent *ent;
    DadosExtraidos lista[100];
    int n = 0;

    while ((ent = readdir(d)) != NULL) {
        if (strstr(ent->d_name, ".pdf") || strstr(ent->d_name, ".PDF")) {
            char caminho[300];
            snprintf(caminho, sizeof(caminho), "data/%s", ent->d_name);
            printf("Processando: %s\n", caminho);
            lista[n++] = processar_laudo(caminho);
        }
    }

    closedir(d);

    gerar_planilha(lista, n);

    return 0;
}