import './App.css';
import ExcelBrandingTool from './ExcelBrandingTool';
import { ThemeProvider, createTheme } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';
import Box from '@mui/material/Box';
import Typography from '@mui/material/Typography';
import Container from '@mui/material/Container';

const theme = createTheme({
  palette: {
    primary: {
      main: '#00796b',
    },
    secondary: {
      main: '#009688',
    },
    background: {
      default: '#fafafa',
      paper: '#ffffff',
    },
  },
  typography: {
    fontFamily: '"Roboto", "Helvetica", "Arial", sans-serif',
  },
});

function App() {
  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Box sx={{ minHeight: '100vh', backgroundColor: '#fafafa' }}>
        {/* Sticky Header */}
        <Box
          sx={{
            backgroundColor: '#ffffff',
            borderBottom: '1px solid #e0e0e0',
            py: 2,
            position: 'sticky',
            top: 0,
            zIndex: 100,
            boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
          }}
        >
          <Container maxWidth="lg">
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
              <Box
                sx={{
                  width: 40,
                  height: 40,
                  borderRadius: '50%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  flexShrink: 0,
                  backgroundImage: `url(${process.env.PUBLIC_URL}/Z_symbol.png)`,
                  backgroundSize: 'contain',
                  backgroundPosition: 'center',
                  backgroundRepeat: 'no-repeat',
                }}
              >
              </Box>
              <Box>
                <Typography
                  variant="subtitle1"
                  sx={{
                    fontWeight: 700,
                    color: '#1a1a1a',
                    letterSpacing: 0.3,
                    lineHeight: 1.2,
                  }}
                >
                  Zutari Excel Template Manager
                </Typography>
                <Typography variant="caption" sx={{ color: '#999' }}>
                  Batch apply templates to Excel files
                </Typography>
              </Box>
            </Box>
          </Container>
        </Box>

        {/* Main Content */}
        <ExcelBrandingTool />
      </Box>
    </ThemeProvider>
  );
}

export default App;
