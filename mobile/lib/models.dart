class PichinchaRecord {
  final DateTime? fecha;
  final String concepto;
  final String documento;
  final String tipo;
  final String beneficiario;
  final double monto;
  final double? saldo;

  PichinchaRecord({
    required this.fecha,
    required this.concepto,
    required this.documento,
    required this.tipo,
    required this.beneficiario,
    required this.monto,
    required this.saldo,
  });
}

class SheetSummary {
  final String sheet;
  final int rowsBefore;
  final int rowsRemoved;
  final int rowsAfter;
  final int transactions;
  final int incomeCount;
  final int expenseCount;
  final double totalIncome;
  final double totalExpenses;
  final double netFlow;

  SheetSummary({
    required this.sheet,
    required this.rowsBefore,
    required this.rowsRemoved,
    required this.rowsAfter,
    required this.transactions,
    required this.incomeCount,
    required this.expenseCount,
    required this.totalIncome,
    required this.totalExpenses,
    required this.netFlow,
  });
}

class ProcessResult {
  final List<PichinchaRecord> records;
  final List<SheetSummary> summaries;
  final List<int> excelBytes;

  ProcessResult({
    required this.records,
    required this.summaries,
    required this.excelBytes,
  });
}
