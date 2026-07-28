import express from 'express';
import cors from 'cors';

const app = express();
app.use(cors());
app.use(express.json());

app.get('/api/health', (req, res) => {
  res.json({ status: 'API is running, ready to accept payroll uploads!' });
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`[Backend] Server listening on http://localhost:${PORT}`);
});
