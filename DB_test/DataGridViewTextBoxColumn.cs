DataGridViewTextBoxColumn box1 = new DataGridViewTextBoxColumn();
            box1.HeaderText = "Text1";
            box1.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dataGridView1.Columns.Add(box1);

            DataGridViewTextBoxColumn box2 = new DataGridViewTextBoxColumn();
            box2.HeaderText = "Text2";
            box2.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dataGridView1.Columns.Add(box2);

            dataGridView1.Rows.Add(5);
            dataGridView1[0, 0].Value = "1";
            dataGridView1[1, 1].Value = "2";