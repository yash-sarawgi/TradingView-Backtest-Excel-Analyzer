#!/usr/bin/env python3
"""
TradingView Backtest Analyzer
Automatically adds comprehensive analysis to TradingView backtest exports

Usage:
    python tradingview_analyzer.py <input_file.xlsx>
    python tradingview_analyzer.py <input_file.xlsx> <output_file.xlsx>
    
For batch processing:
    python tradingview_analyzer.py --batch <directory>
"""

""""
By - Yash Sarawgi, CFTe, CMT L3
"""


import sys
import os
import pandas as pd
import json
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference

def analyze_trades(trades_df):
    """Calculate all trading metrics from trades dataframe"""
    
    # Filter only Exit trades (each complete trade)
    exit_trades = trades_df[trades_df['Type'].str.contains('Exit', case=False, na=False)].copy()
    
    if len(exit_trades) == 0:
        print("Warning: No exit trades found in data")
        return None, None
    
    # Add time-based columns
    exit_trades['Day of Week'] = exit_trades['Date and time'].dt.day_name()
    exit_trades['Hour'] = exit_trades['Date and time'].dt.hour
    exit_trades['Month'] = exit_trades['Date and time'].dt.month_name()
    exit_trades['Year'] = exit_trades['Date and time'].dt.year
    exit_trades['Date'] = exit_trades['Date and time'].dt.date
    
    # Calculate basic metrics
    total_trades = len(exit_trades)
    winning_trades = len(exit_trades[exit_trades['Net P&L USD'] > 0])
    losing_trades = len(exit_trades[exit_trades['Net P&L USD'] < 0])
    win_rate = (winning_trades / total_trades * 100) if total_trades > 0 else 0
    
    total_profit = exit_trades[exit_trades['Net P&L USD'] > 0]['Net P&L USD'].sum()
    total_loss = abs(exit_trades[exit_trades['Net P&L USD'] < 0]['Net P&L USD'].sum())
    profit_factor = (total_profit / total_loss) if total_loss > 0 else 0
    
    avg_win = exit_trades[exit_trades['Net P&L USD'] > 0]['Net P&L USD'].mean() if winning_trades > 0 else 0
    avg_loss = exit_trades[exit_trades['Net P&L USD'] < 0]['Net P&L USD'].mean() if losing_trades > 0 else 0
    avg_trade = exit_trades['Net P&L USD'].mean()
    
    largest_win = exit_trades['Net P&L USD'].max()
    largest_loss = exit_trades['Net P&L USD'].min()
    
    # Consecutive wins/losses
    max_consecutive_wins = 0
    max_consecutive_losses = 0
    current_streak_wins = 0
    current_streak_losses = 0
    
    for pnl in exit_trades['Net P&L USD']:
        if pnl > 0:
            current_streak_wins += 1
            current_streak_losses = 0
            max_consecutive_wins = max(max_consecutive_wins, current_streak_wins)
        else:
            current_streak_losses += 1
            current_streak_wins = 0
            max_consecutive_losses = max(max_consecutive_losses, current_streak_losses)
    
    # Day of week analysis
    dow_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    dow_stats = exit_trades.groupby('Day of Week').agg({
        'Net P&L USD': ['sum', 'count', 'mean'],
    }).round(2)
    dow_stats.columns = ['Total P&L', 'Trades', 'Avg P&L']
    dow_stats['Win Rate %'] = exit_trades.groupby('Day of Week').apply(
        lambda x: (len(x[x['Net P&L USD'] > 0]) / len(x) * 100) if len(x) > 0 else 0
    ).round(2)
    dow_stats = dow_stats.reindex([d for d in dow_order if d in dow_stats.index])
    
    # Hour analysis
    hour_stats = exit_trades.groupby('Hour').agg({
        'Net P&L USD': ['sum', 'count', 'mean'],
    }).round(2)
    hour_stats.columns = ['Total P&L', 'Trades', 'Avg P&L']
    hour_stats['Win Rate %'] = exit_trades.groupby('Hour').apply(
        lambda x: (len(x[x['Net P&L USD'] > 0]) / len(x) * 100) if len(x) > 0 else 0
    ).round(2)
    
    # Month analysis
    month_stats = exit_trades.groupby('Month').agg({
        'Net P&L USD': ['sum', 'count', 'mean'],
    }).round(2)
    month_stats.columns = ['Total P&L', 'Trades', 'Avg P&L']
    month_stats['Win Rate %'] = exit_trades.groupby('Month').apply(
        lambda x: (len(x[x['Net P&L USD'] > 0]) / len(x) * 100) if len(x) > 0 else 0
    ).round(2)
    
    # Daily equity curve
    daily_pnl = exit_trades.groupby('Date')['Net P&L USD'].sum().reset_index()
    daily_pnl['Cumulative P&L'] = daily_pnl['Net P&L USD'].cumsum()
    
    # Drawdown analysis
    exit_trades_sorted = exit_trades.sort_values('Date and time')
    cumulative_pnl = exit_trades_sorted['Net P&L USD'].cumsum()
    running_max = cumulative_pnl.expanding().max()
    drawdown = cumulative_pnl - running_max
    max_drawdown = drawdown.min()
    max_drawdown_pct = (max_drawdown / running_max[drawdown.idxmin()]) * 100 if running_max[drawdown.idxmin()] != 0 else 0
    
    # Risk metrics
    returns = exit_trades['Net P&L %'].values
    avg_return = returns.mean()
    std_return = returns.std()
    sharpe_like = (avg_return / std_return) if std_return != 0 else 0
    
    expectancy = (exit_trades[exit_trades['Net P&L USD'] > 0]['Net P&L USD'].mean() * 
                  (exit_trades[exit_trades['Net P&L USD'] > 0].shape[0] / exit_trades.shape[0])) + \
                 (exit_trades[exit_trades['Net P&L USD'] < 0]['Net P&L USD'].mean() * 
                  (exit_trades[exit_trades['Net P&L USD'] < 0].shape[0] / exit_trades.shape[0]))
    
    # Trade distribution
    distribution = [
        ('< -$500', len(exit_trades[exit_trades['Net P&L USD'] < -500])),
        ('-$500 to -$250', len(exit_trades[(exit_trades['Net P&L USD'] >= -500) & (exit_trades['Net P&L USD'] < -250)])),
        ('-$250 to -$100', len(exit_trades[(exit_trades['Net P&L USD'] >= -250) & (exit_trades['Net P&L USD'] < -100)])),
        ('-$100 to $0', len(exit_trades[(exit_trades['Net P&L USD'] >= -100) & (exit_trades['Net P&L USD'] < 0)])),
        ('$0 to $100', len(exit_trades[(exit_trades['Net P&L USD'] >= 0) & (exit_trades['Net P&L USD'] < 100)])),
        ('$100 to $250', len(exit_trades[(exit_trades['Net P&L USD'] >= 100) & (exit_trades['Net P&L USD'] < 250)])),
        ('$250 to $500', len(exit_trades[(exit_trades['Net P&L USD'] >= 250) & (exit_trades['Net P&L USD'] < 500)])),
        ('> $500', len(exit_trades[exit_trades['Net P&L USD'] >= 500])),
    ]
    
    metrics = {
        'total_trades': int(total_trades),
        'winning_trades': int(winning_trades),
        'losing_trades': int(losing_trades),
        'win_rate': float(win_rate),
        'total_profit': float(total_profit),
        'total_loss': float(total_loss),
        'profit_factor': float(profit_factor),
        'avg_win': float(avg_win),
        'avg_loss': float(avg_loss),
        'avg_trade': float(avg_trade),
        'largest_win': float(largest_win),
        'largest_loss': float(largest_loss),
        'max_consecutive_wins': int(max_consecutive_wins),
        'max_consecutive_losses': int(max_consecutive_losses),
        'max_drawdown': float(max_drawdown),
        'max_drawdown_pct': float(max_drawdown_pct),
        'sharpe_like': float(sharpe_like),
        'expectancy': float(expectancy),
        'std_return': float(std_return),
        'avg_return': float(avg_return),
    }
    
    data_frames = {
        'dow_stats': dow_stats,
        'hour_stats': hour_stats,
        'month_stats': month_stats,
        'daily_pnl': daily_pnl,
        'distribution': distribution,
    }
    
    return metrics, data_frames

def create_analysis_sheet(wb, metrics, data_frames):
    """Create the Advanced Analysis sheet with all metrics and charts"""
    
    # Create new sheet
    if 'Advanced Analysis' in wb.sheetnames:
        del wb['Advanced Analysis']
    ws = wb.create_sheet('Advanced Analysis', 0)
    
    # Styling
    header_font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    subheader_font = Font(name='Arial', size=11, bold=True)
    subheader_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    metric_font = Font(name='Arial', size=10)
    
    # Title
    ws['A1'] = 'TradingView Backtest - Advanced Analysis'
    ws['A1'].font = Font(name='Arial', size=14, bold=True)
    ws.merge_cells('A1:F1')
    ws['A1'].alignment = Alignment(horizontal='center')
    
    # === KEY METRICS SECTION ===
    row = 3
    ws[f'A{row}'] = 'KEY PERFORMANCE METRICS'
    ws[f'A{row}'].font = header_font
    ws[f'A{row}'].fill = header_fill
    ws.merge_cells(f'A{row}:B{row}')
    
    row += 1
    metrics_data = [
        ('Total Trades', metrics['total_trades'], ''),
        ('Winning Trades', metrics['winning_trades'], ''),
        ('Losing Trades', metrics['losing_trades'], ''),
        ('Win Rate', metrics['win_rate'], '%'),
        ('', '', ''),
        ('Total Profit', metrics['total_profit'], '$'),
        ('Total Loss', abs(metrics['total_loss']), '$'),
        ('Net Profit', metrics['total_profit'] - abs(metrics['total_loss']), '$'),
        ('Profit Factor', metrics['profit_factor'], ''),
        ('', '', ''),
        ('Average Win', metrics['avg_win'], '$'),
        ('Average Loss', abs(metrics['avg_loss']), '$'),
        ('Average Trade', metrics['avg_trade'], '$'),
        ('', '', ''),
        ('Largest Win', metrics['largest_win'], '$'),
        ('Largest Loss', abs(metrics['largest_loss']), '$'),
        ('', '', ''),
        ('Max Consecutive Wins', metrics['max_consecutive_wins'], ''),
        ('Max Consecutive Losses', metrics['max_consecutive_losses'], ''),
    ]
    
    for metric_name, value, suffix in metrics_data:
        ws[f'A{row}'] = metric_name
        ws[f'A{row}'].font = metric_font
        if metric_name:
            ws[f'B{row}'] = value
            ws[f'B{row}'].font = Font(name='Arial', size=10, bold=True)
            if suffix == '$':
                ws[f'B{row}'].number_format = '$#,##0.00'
            elif suffix == '%':
                ws[f'B{row}'].number_format = '0.00"%"'
            else:
                ws[f'B{row}'].number_format = '#,##0.00'
        row += 1
    
    # === DAY OF WEEK ANALYSIS ===
    ws[f'D{3}'] = 'DAY OF WEEK ANALYSIS'
    ws[f'D{3}'].font = header_font
    ws[f'D{3}'].fill = header_fill
    ws.merge_cells(f'D{3}:G{3}')
    
    row = 4
    headers = ['Day', 'Total P&L', 'Trades', 'Avg P&L', 'Win Rate %']
    for col_idx, header in enumerate(headers, start=4):
        cell = ws.cell(row=row, column=col_idx)
        cell.value = header
        cell.font = subheader_font
        cell.fill = subheader_fill
        cell.alignment = Alignment(horizontal='center')
    
    row += 1
    dow_start_row = row
    for idx, data_row in data_frames['dow_stats'].iterrows():
        ws.cell(row=row, column=4, value=idx)
        ws.cell(row=row, column=5, value=data_row['Total P&L']).number_format = '$#,##0.00'
        ws.cell(row=row, column=6, value=data_row['Trades']).number_format = '#,##0'
        ws.cell(row=row, column=7, value=data_row['Avg P&L']).number_format = '$#,##0.00'
        ws.cell(row=row, column=8, value=data_row['Win Rate %']).number_format = '0.00"%"'
        row += 1
    dow_end_row = row - 1
    
    # === HOUR ANALYSIS ===
    row += 1
    ws[f'D{row}'] = 'HOUR OF DAY ANALYSIS'
    ws[f'D{row}'].font = header_font
    ws[f'D{row}'].fill = header_fill
    ws.merge_cells(f'D{row}:G{row}')
    
    row += 1
    headers = ['Hour', 'Total P&L', 'Trades', 'Avg P&L', 'Win Rate %']
    for col_idx, header in enumerate(headers, start=4):
        cell = ws.cell(row=row, column=col_idx)
        cell.value = header
        cell.font = subheader_font
        cell.fill = subheader_fill
        cell.alignment = Alignment(horizontal='center')
    
    row += 1
    hour_start_row = row
    for idx, data_row in data_frames['hour_stats'].iterrows():
        ws.cell(row=row, column=4, value=int(idx))
        ws.cell(row=row, column=5, value=data_row['Total P&L']).number_format = '$#,##0.00'
        ws.cell(row=row, column=6, value=data_row['Trades']).number_format = '#,##0'
        ws.cell(row=row, column=7, value=data_row['Avg P&L']).number_format = '$#,##0.00'
        ws.cell(row=row, column=8, value=data_row['Win Rate %']).number_format = '0.00"%"'
        row += 1
    hour_end_row = row - 1
    
    # === MONTH ANALYSIS ===
    row += 1
    ws[f'D{row}'] = 'MONTH ANALYSIS'
    ws[f'D{row}'].font = header_font
    ws[f'D{row}'].fill = header_fill
    ws.merge_cells(f'D{row}:G{row}')
    
    row += 1
    headers = ['Month', 'Total P&L', 'Trades', 'Avg P&L', 'Win Rate %']
    for col_idx, header in enumerate(headers, start=4):
        cell = ws.cell(row=row, column=col_idx)
        cell.value = header
        cell.font = subheader_font
        cell.fill = subheader_fill
        cell.alignment = Alignment(horizontal='center')
    
    row += 1
    for idx, data_row in data_frames['month_stats'].iterrows():
        ws.cell(row=row, column=4, value=idx)
        ws.cell(row=row, column=5, value=data_row['Total P&L']).number_format = '$#,##0.00'
        ws.cell(row=row, column=6, value=data_row['Trades']).number_format = '#,##0'
        ws.cell(row=row, column=7, value=data_row['Avg P&L']).number_format = '$#,##0.00'
        ws.cell(row=row, column=8, value=data_row['Win Rate %']).number_format = '0.00"%"'
        row += 1
    
    # === TRADE DISTRIBUTION ===
    dist_row = 27
    ws[f'J{dist_row}'] = 'TRADE P&L DISTRIBUTION'
    ws[f'J{dist_row}'].font = header_font
    ws[f'J{dist_row}'].fill = header_fill
    ws.merge_cells(f'J{dist_row}:K{dist_row}')
    
    dist_row += 1
    ws.cell(row=dist_row, column=10, value='P&L Range').font = subheader_font
    ws.cell(row=dist_row, column=11, value='Count').font = subheader_font
    ws.cell(row=dist_row, column=12, value='%').font = subheader_font
    
    dist_row += 1
    for range_label, count in data_frames['distribution']:
        ws.cell(row=dist_row, column=10, value=range_label)
        ws.cell(row=dist_row, column=11, value=count)
        ws.cell(row=dist_row, column=12, value=count/metrics['total_trades']*100).number_format = '0.00"%"'
        dist_row += 1
    
    # === DRAWDOWN & RISK METRICS ===
    risk_row = dist_row + 2
    ws[f'J{risk_row}'] = 'DRAWDOWN & RISK METRICS'
    ws[f'J{risk_row}'].font = header_font
    ws[f'J{risk_row}'].fill = header_fill
    ws.merge_cells(f'J{risk_row}:K{risk_row}')
    
    risk_row += 1
    risk_metrics = [
        ('Max Drawdown ($)', abs(metrics['max_drawdown']), '$'),
        ('Max Drawdown (%)', abs(metrics['max_drawdown_pct']), '%'),
        ('Sharpe-like Ratio', metrics['sharpe_like'], ''),
        ('Expectancy per Trade', metrics['expectancy'], '$'),
        ('Std Dev of Returns', metrics['std_return'], '%'),
        ('Avg Return per Trade', metrics['avg_return'], '%'),
    ]
    
    for metric_name, value, suffix in risk_metrics:
        ws.cell(row=risk_row, column=10, value=metric_name)
        ws.cell(row=risk_row, column=11, value=value)
        if suffix == '$':
            ws.cell(row=risk_row, column=11).number_format = '$#,##0.00'
        elif suffix == '%':
            ws.cell(row=risk_row, column=11).number_format = '0.00"%"'
        else:
            ws.cell(row=risk_row, column=11).number_format = '#,##0.00'
        risk_row += 1
    
    # === DAILY EQUITY CURVE ===
    equity_row = row + 2
    ws[f'A{equity_row}'] = 'DAILY EQUITY CURVE'
    ws[f'A{equity_row}'].font = header_font
    ws[f'A{equity_row}'].fill = header_fill
    ws.merge_cells(f'A{equity_row}:C{equity_row}')
    
    equity_row += 1
    ws.cell(row=equity_row, column=1, value='Date').font = subheader_font
    ws.cell(row=equity_row, column=1).fill = subheader_fill
    ws.cell(row=equity_row, column=2, value='Daily P&L').font = subheader_font
    ws.cell(row=equity_row, column=2).fill = subheader_fill
    ws.cell(row=equity_row, column=3, value='Cumulative P&L').font = subheader_font
    ws.cell(row=equity_row, column=3).fill = subheader_fill
    
    equity_row += 1
    equity_start_row = equity_row
    for _, data_row in data_frames['daily_pnl'].iterrows():
        ws.cell(row=equity_row, column=1, value=str(data_row['Date']))
        ws.cell(row=equity_row, column=2, value=data_row['Net P&L USD']).number_format = '$#,##0.00'
        ws.cell(row=equity_row, column=3, value=data_row['Cumulative P&L']).number_format = '$#,##0.00'
        equity_row += 1
    equity_end_row = equity_row - 1
    
    # === ADD CHARTS ===
    # Chart 1: Day of Week P&L
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "P&L by Day of Week"
    chart1.y_axis.title = 'Total P&L ($)'
    chart1.x_axis.title = 'Day of Week'
    data = Reference(ws, min_col=5, min_row=4, max_row=dow_end_row, max_col=5)
    cats = Reference(ws, min_col=4, min_row=5, max_row=dow_end_row)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.height = 10
    chart1.width = 20
    ws.add_chart(chart1, "J3")
    
    # Chart 2: Hour Performance
    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 11
    chart2.title = "P&L by Hour of Day"
    chart2.y_axis.title = 'Total P&L ($)'
    chart2.x_axis.title = 'Hour'
    hour_header_row = hour_start_row - 1
    data = Reference(ws, min_col=5, min_row=hour_header_row, max_row=hour_end_row, max_col=5)
    cats = Reference(ws, min_col=4, min_row=hour_start_row, max_row=hour_end_row)
    chart2.add_data(data, titles_from_data=True)
    chart2.set_categories(cats)
    chart2.height = 10
    chart2.width = 20
    ws.add_chart(chart2, "J20")
    
    # Chart 3: Win Rate by Day
    chart3 = BarChart()
    chart3.type = "col"
    chart3.style = 12
    chart3.title = "Win Rate by Day of Week"
    chart3.y_axis.title = 'Win Rate (%)'
    chart3.x_axis.title = 'Day of Week'
    data = Reference(ws, min_col=8, min_row=4, max_row=dow_end_row, max_col=8)
    cats = Reference(ws, min_col=4, min_row=5, max_row=dow_end_row)
    chart3.add_data(data, titles_from_data=True)
    chart3.set_categories(cats)
    chart3.height = 10
    chart3.width = 20
    ws.add_chart(chart3, "T3")
    
    # Chart 4: Equity Curve
    chart4 = LineChart()
    chart4.style = 13
    chart4.title = "Cumulative P&L Equity Curve"
    chart4.y_axis.title = 'Cumulative P&L ($)'
    chart4.x_axis.title = 'Trading Days'
    equity_header_row = equity_start_row - 1
    data = Reference(ws, min_col=3, min_row=equity_header_row, max_row=min(equity_end_row, equity_start_row + 200), max_col=3)
    chart4.add_data(data, titles_from_data=True)
    chart4.height = 12
    chart4.width = 25
    ws.add_chart(chart4, "J37")
    
    # Set column widths
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['J'].width = 25
    ws.column_dimensions['K'].width = 15
    ws.column_dimensions['L'].width = 12

def process_file(input_file, output_file=None):
    """Process a single TradingView backtest file"""
    
    if not os.path.exists(input_file):
        print(f"Error: File not found: {input_file}")
        return False
    
    print(f"Processing: {input_file}")
    
    try:
        # Load trades data
        trades_df = pd.read_excel(input_file, sheet_name='List of trades')
        print(f"  Loaded {len(trades_df)} trade records")
        
        # Analyze trades
        metrics, data_frames = analyze_trades(trades_df)
        
        if metrics is None:
            print("  Error: No valid trade data found")
            return False
        
        print(f"  Analyzed {metrics['total_trades']} trades")
        print(f"  Win Rate: {metrics['win_rate']:.2f}%")
        print(f"  Profit Factor: {metrics['profit_factor']:.2f}")
        
        # Load workbook and create analysis sheet
        wb = load_workbook(input_file)
        create_analysis_sheet(wb, metrics, data_frames)
        
        # Save
        if output_file is None:
            base, ext = os.path.splitext(input_file)
            output_file = f"{base}_analyzed{ext}"
        
        wb.save(output_file)
        print(f"  ✓ Saved to: {output_file}")
        
        return True
        
    except Exception as e:
        print(f"  Error processing file: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def batch_process(directory):
    """Process all Excel files in a directory"""
    
    if not os.path.isdir(directory):
        print(f"Error: Directory not found: {directory}")
        return
    
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx') and not f.startswith('~')]
    
    if not files:
        print(f"No Excel files found in {directory}")
        return
    
    print(f"Found {len(files)} files to process\n")
    
    success_count = 0
    for i, filename in enumerate(files, 1):
        input_path = os.path.join(directory, filename)
        print(f"[{i}/{len(files)}] {filename}")
        
        if process_file(input_path):
            success_count += 1
        print()
    
    print(f"Batch complete: {success_count}/{len(files)} files processed successfully")

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    
    if sys.argv[1] == '--batch':
        if len(sys.argv) < 3:
            print("Error: Please specify a directory for batch processing")
            sys.exit(1)
        batch_process(sys.argv[2])
    else:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        
        if process_file(input_file, output_file):
            print("\n✓ Analysis complete!")
        else:
            print("\n✗ Analysis failed")
            sys.exit(1)

if __name__ == '__main__':
    main()
