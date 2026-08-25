import 'dart:typed_data';
import 'package:excel/excel.dart';
import 'package:intl/intl.dart';
import 'models.dart';

class PichinchaMobileEngine {
  static const List<String> cleanHeaders = [
    'Fecha',
    'Concepto',
    'Nro. Documento',
    'Tipo',
    'Beneficiario',
    'Monto',
    'Saldo',
  ];

  static String _cleanText(dynamic val) {
    if (val == null) return '';
    return val.toString().replaceAll('_x000D_', ' ').replaceAll(RegExp(r'\s+'), ' ').trim();
  }

  static String _normalize(dynamic val) {
    String text = _cleanText(val).toUpperCase();
    return text.replaceAll(RegExp(r'[^A-Z0-9]+'), '');
  }

  static double? _parseMoney(dynamic val) {
    if (val == null) return 0.0;
    if (val is num) return val.toDouble();
    String text = _cleanText(val).replaceAll('\$', '').replaceAll(' ', '');
    if (text.isEmpty) return 0.0;

    bool negative = text.startsWith('-') || (text.startsWith('(') && text.endsWith(')')) || text.endsWith('-');
    text = text.replaceAll(RegExp(r'[()\-+]'), '');

    if (text.contains(',') && text.contains('.')) {
      text = text.lastIndexOf(',') > text.lastIndexOf('.')
          ? text.replaceAll('.', '').replaceAll(',', '.')
          : text.replaceAll(',', '');
    } else if (RegExp(r',\d{1,2}$').hasMatch(text)) {
      text = text.replaceAll('.', '').replaceAll(',', '.');
    } else {
      text = text.replaceAll(',', '');
    }

    double? parsed = double.tryParse(text);
    if (parsed == null) return 0.0;
    return negative ? -parsed : parsed;
  }

  static DateTime? _parseDate(dynamic val) {
    if (val == null || val == '') return null;
    if (val is DateTime) return val;
    String text = _cleanText(val).replaceAll(',', ' ');

    final isoMatch = RegExp(
      r'^(?<year>\d{4})-(?<month>\d{1,2})-(?<day>\d{1,2})(?:\s+(?<hour>\d{1,2}):(?<minute>\d{2})(?:\s*(?<ampm>AM|PM))?)?',
      caseSensitive: false,
    ).firstMatch(text);

    if (isoMatch != null) {
      try {
        int y = int.parse(isoMatch.namedGroup('year')!);
        int m = int.parse(isoMatch.namedGroup('month')!);
        int d = int.parse(isoMatch.namedGroup('day')!);
        int h = int.parse(isoMatch.namedGroup('hour') ?? '0');
        int min = int.parse(isoMatch.namedGroup('minute') ?? '0');
        String? ampm = isoMatch.namedGroup('ampm');
        if (ampm != null) {
          h = (h % 12) + (ampm.toUpperCase() == 'PM' ? 12 : 0);
        }
        return DateTime(y, m, d, h, min);
      } catch (_) {}
    }

    final formats = [
      'dd/MM/yyyy HH:mm',
      'dd/MM/yyyy hh:mm a',
      'dd/MM/yyyy',
      'dd-MM-yyyy HH:mm',
      'dd-MM-yyyy hh:mm a',
      'dd-MM-yyyy',
      'yyyy-MM-dd HH:mm',
      'yyyy-MM-dd',
    ];

    for (var f in formats) {
      try {
        return DateFormat(f).parse(text);
      } catch (_) {}
    }
    return null;
  }

  static ProcessResult processExcelBytes(Uint8List inputBytes) {
    var excel = Excel.decodeBytes(inputBytes);
    List<PichinchaRecord> allRecords = [];
    List<SheetSummary> summaries = [];

    for (var table in excel.tables.keys) {
      var rows = excel.tables[table]!.rows;
      if (rows.isEmpty) continue;

      int? headerRowIndex;
      for (int i = 0; i < rows.length; i++) {
        var row = rows[i];
        bool hasFecha = row.any((c) => _normalize(c?.value) == 'FECHA');
        bool hasConcepto = row.any((c) => _normalize(c?.value) == 'CONCEPTO' || _normalize(c?.value) == 'DESCRIPCION');
        if (hasFecha && hasConcepto) {
          headerRowIndex = i;
          break;
        }
      }

      if (headerRowIndex == null) continue;
      var headerRow = rows[headerRowIndex];

      int? getColIdx(List<String> labels) {
        var wanted = labels.map(_normalize).toSet();
        for (int c = 0; c < headerRow.length; c++) {
          if (wanted.contains(_normalize(headerRow[c]?.value))) return c;
        }
        return null;
      }

      int? dateCol = getColIdx(['Fecha']);
      int? descCol = getColIdx(['Concepto', 'Descripcion']);
      int? docCol = getColIdx(['Nro. Documento', 'Documento', 'Nro Documento']);
      int? typeCol = getColIdx(['Tipo']);
      int? benCol = getColIdx(['Beneficiario']);
      int? amountCol = getColIdx(['Monto', 'Valor']);
      int? balanceCol = getColIdx(['Saldo']);

      String lookAhead(int rIdx, int? cIdx, String kind) {
        if (cIdx == null) return '';
        List<int> candidateCols = [cIdx];
        if (kind == 'doc') {
          if (cIdx > 0) candidateCols.add(cIdx - 1);
          candidateCols.add(cIdx + 1);
        }
        for (int offset = 0; offset < 3; offset++) {
          int targetR = rIdx + offset;
          if (targetR >= rows.length) break;
          var r = rows[targetR];
          for (var c in candidateCols) {
            if (c >= r.length) continue;
            String val = _cleanText(r[c]?.value);
            if (val.isEmpty || val.toLowerCase().startsWith('total')) continue;
            if (kind == 'doc' && RegExp(r'^\d{5,}$').hasMatch(val)) return val;
            if (kind != 'doc' && !['Monto', 'Saldo', 'Tipo'].contains(val)) return val;
          }
        }
        return '';
      }

      List<PichinchaRecord> sheetRecords = [];
      int incomeCount = 0;
      int expenseCount = 0;
      double totalIncome = 0.0;
      double totalExpenses = 0.0;

      for (int r = headerRowIndex + 1; r < rows.length; r++) {
        var row = rows[r];
        if (dateCol == null || dateCol >= row.length || row[dateCol]?.value == null) continue;
        if (descCol == null || descCol >= row.length || row[descCol]?.value == null) continue;

        String typ = typeCol != null && typeCol < row.length ? _cleanText(row[typeCol]?.value).toLowerCase() : '';
        double? amount = amountCol != null && amountCol < row.length ? _parseMoney(row[amountCol]?.value) : null;
        if (amount == null || amount == 0.0) continue;

        DateTime? dt = _parseDate(row[dateCol]?.value);
        String desc = _cleanText(row[descCol]?.value);
        String doc = lookAhead(r, docCol, 'doc');
        String ben = lookAhead(r, benCol, 'ben');
        double? balance = balanceCol != null && balanceCol < row.length ? _parseMoney(row[balanceCol]?.value) : null;

        String movementType;
        double finalAmount;

        if (typ.contains('cred')) {
          movementType = 'Crédito';
          finalAmount = amount.abs();
          incomeCount++;
          totalIncome += finalAmount;
        } else if (typ.contains('deb')) {
          movementType = 'Débito';
          finalAmount = -amount.abs();
          expenseCount++;
          totalExpenses += amount.abs();
        } else {
          if (amount > 0) {
            movementType = 'Crédito';
            finalAmount = amount;
            incomeCount++;
            totalIncome += finalAmount;
          } else {
            movementType = 'Débito';
            finalAmount = amount;
            expenseCount++;
            totalExpenses += amount.abs();
          }
        }

        var record = PichinchaRecord(
          fecha: dt,
          concepto: desc,
          documento: doc,
          tipo: movementType,
          beneficiario: ben,
          monto: finalAmount,
          saldo: balance,
        );
        sheetRecords.add(record);
        allRecords.add(record);
      }

      int outRows = sheetRecords.length + 1;
      summaries.add(SheetSummary(
        sheet: table,
        rowsBefore: rows.length,
        rowsRemoved: (rows.length - outRows) > 0 ? rows.length - outRows : 0,
        rowsAfter: outRows,
        transactions: sheetRecords.length,
        incomeCount: incomeCount,
        expenseCount: expenseCount,
        totalIncome: totalIncome,
        totalExpenses: totalExpenses,
        netFlow: totalIncome - totalExpenses,
      ));
    }

    if (allRecords.isEmpty) {
      throw Exception('No se encontraron movimientos válidos de Banco Pichincha en el archivo.');
    }

    // Generate Clean Output Excel
    var outExcel = Excel.createExcel();
    String sheetName = 'Movimientos';
    outExcel.rename(outExcel.getDefaultSheet()!, sheetName);
    Sheet sheet = outExcel[sheetName];

    // Headers
    sheet.appendRow(cleanHeaders.map((h) => TextCellValue(h)).toList());

    // Date Formatter: dd-mm-yyyy hh:mm
    final dateFmt = DateFormat('dd-MM-yyyy HH:mm');
    final moneyFmt = NumberFormat('\$#,##0.00;-\$#,##0.00;0.00');

    for (var rec in allRecords) {
      sheet.appendRow([
        TextCellValue(rec.fecha != null ? dateFmt.format(rec.fecha!) : ''),
        TextCellValue(rec.concepto),
        TextCellValue(rec.documento),
        TextCellValue(rec.tipo),
        TextCellValue(rec.beneficiario),
        DoubleCellValue(rec.monto),
        rec.saldo != null ? DoubleCellValue(rec.saldo!) : TextCellValue(''),
      ]);
    }

    var encodedBytes = outExcel.save()!;
    return ProcessResult(
      records: allRecords,
      summaries: summaries,
      excelBytes: encodedBytes,
    );
  }
}
