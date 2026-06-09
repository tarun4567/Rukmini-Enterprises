document.addEventListener('DOMContentLoaded', () => {
    // 1. Sales Trend Line Chart
    const salesChartCanvas = document.getElementById('salesTrendChart');
    if (salesChartCanvas) {
        // Retrieve dynamic data injected via data attributes
        const labels = JSON.parse(salesChartCanvas.getAttribute('data-labels') || '[]');
        const data = JSON.parse(salesChartCanvas.getAttribute('data-values') || '[]');
        
        const ctx = salesChartCanvas.getContext('2d');
        new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Sales Revenue (₹)',
                    data: data,
                    borderColor: '#2563eb',
                    backgroundColor: 'rgba(37, 99, 235, 0.05)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.35,
                    pointBackgroundColor: '#2563eb',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 2,
                    pointRadius: 5,
                    pointHoverRadius: 7
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        padding: 12,
                        cornerRadius: 8,
                        fontFamily: 'Plus Jakarta Sans',
                        callbacks: {
                            label: function(context) {
                                return ' Sales: ₹' + context.raw.toLocaleString();
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            font: {
                                family: 'Plus Jakarta Sans',
                                size: 11
                            },
                            color: '#64748b'
                        }
                    },
                    y: {
                        grid: {
                            color: '#e2e8f0',
                            drawBorder: false
                        },
                        ticks: {
                            font: {
                                family: 'Plus Jakarta Sans',
                                size: 11
                            },
                            color: '#64748b',
                            callback: function(value) {
                                return '₹' + value.toLocaleString();
                            }
                        }
                    }
                }
            }
        });
    }

    // 2. Category Stock Value Doughnut Chart
    const categoryChartCanvas = document.getElementById('categoryStockChart');
    if (categoryChartCanvas) {
        const labels = JSON.parse(categoryChartCanvas.getAttribute('data-labels') || '[]');
        const data = JSON.parse(categoryChartCanvas.getAttribute('data-values') || '[]');
        
        const ctx = categoryChartCanvas.getContext('2d');
        new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: labels,
                datasets: [{
                    data: data,
                    backgroundColor: [
                        '#2563eb', // Blue
                        '#8b5cf6', // Purple
                        '#10b981', // Green
                        '#f59e0b', // Amber
                        '#ef4444', // Red
                        '#06b6d4'  // Cyan
                    ],
                    borderWidth: 2,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            boxWidth: 12,
                            padding: 16,
                            font: {
                                family: 'Plus Jakarta Sans',
                                size: 11
                            },
                            color: '#64748b'
                        }
                    },
                    tooltip: {
                        padding: 12,
                        cornerRadius: 8,
                        fontFamily: 'Plus Jakarta Sans',
                        callbacks: {
                            label: function(context) {
                                return ' Value: ₹' + context.raw.toLocaleString();
                            }
                        }
                    }
                },
                cutout: '70%'
            }
        });
    }
});
